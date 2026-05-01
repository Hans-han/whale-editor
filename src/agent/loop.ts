import type {
  LLMProvider,
  OfficeManifest,
  AgentConfig,
  AgentResult,
  IterationLog,
  UsageMetrics,
} from '../types/index.js';
import { buildPrompt, appendIterationContext, TOOL_SPECS, PATCH_SCHEMA, VALIDATION_RULES } from './promptBuilder.js';
import { AVAILABLE_TOOLS } from './toolRegistry.js';
import { validatePatch } from '../patches/patchValidator.js';
import { executePatch } from '../patches/patchExecutor.js';
import { generateDocxManifest } from '../office/docx/docxManifest.js';
import { generatePptxManifest } from '../office/pptx/pptxManifest.js';
import { validateDocx } from '../office/docx/docxValidation.js';
import { validatePptx } from '../office/pptx/pptxValidation.js';
import { logger } from '../utils/logger.js';

const DEFAULT_CONFIG: AgentConfig = {
  maxIterations: 4,
  maxPatchOperations: 20,
  maxOutputTokens: 4096,
  model: process.env.DEEPSEEK_MODEL ?? 'deepseek-v4',
  baseUrl: process.env.DEEPSEEK_BASE_URL ?? 'https://api.deepseek.com/v1',
  apiKey: process.env.DEEPSEEK_API_KEY ?? '',
};

export async function runAgent(
  provider: LLMProvider,
  fileBuffer: Buffer,
  fileType: 'docx' | 'pptx',
  userTask: string,
  config: Partial<AgentConfig> = {}
): Promise<AgentResult> {
  const fullConfig = { ...DEFAULT_CONFIG, ...config };
  const iterations: IterationLog[] = [];
  let currentBuffer = fileBuffer;
  let totalUsage: UsageMetrics = {
    inputTokens: 0,
    outputTokens: 0,
    totalTokens: 0,
  };

  const manifest = fileType === 'docx'
    ? await generateDocxManifest(fileBuffer)
    : await generatePptxManifest(fileBuffer);

  const promptResult = buildPrompt({
    systemRules: '',
    toolSpecs: TOOL_SPECS,
    patchSchema: PATCH_SCHEMA,
    validationRules: VALIDATION_RULES,
    manifest,
    userTask,
  });

  logger.info('Starting agent loop', {
    fileType,
    maxIterations: fullConfig.maxIterations,
    stablePrefixHash: promptResult.stablePrefixHash,
  });

  let messages = promptResult.messages;

  for (let i = 0; i < fullConfig.maxIterations; i++) {
    logger.info(`Iteration ${i + 1}/${fullConfig.maxIterations}`);

    const iteration: IterationLog = { iteration: i + 1 };

    try {
      const result = await provider.generateToolCall(messages, AVAILABLE_TOOLS, {
        maxTokens: fullConfig.maxOutputTokens,
        temperature: 0.1,
      });

      totalUsage.inputTokens += result.usage.inputTokens;
      totalUsage.outputTokens += result.usage.outputTokens;
      totalUsage.totalTokens += result.usage.totalTokens;
      if (result.usage.cacheHitTokens !== undefined) {
        totalUsage.cacheHitTokens = (totalUsage.cacheHitTokens ?? 0) + result.usage.cacheHitTokens;
      }
      if (result.usage.cacheMissTokens !== undefined) {
        totalUsage.cacheMissTokens = (totalUsage.cacheMissTokens ?? 0) + result.usage.cacheMissTokens;
      }

      if (!result.toolCalls || result.toolCalls.length === 0) {
        logger.info('No tool calls returned, agent finished');
        iterations.push(iteration);
        break;
      }

      const toolCall = result.toolCalls[0];
      if (toolCall.function.name !== 'apply_patch') {
        throw new Error(`Unexpected tool call: ${toolCall.function.name}`);
      }

      const patchData = JSON.parse(toolCall.function.arguments);
      const validatedPatch = validatePatch(patchData.patch);
      iteration.patchPlan = validatedPatch;

      logger.info('Executing patch', {
        operations: validatedPatch.operations.length,
        intent: validatedPatch.intent,
      });

      const execResult = await executePatch(currentBuffer, validatedPatch);
      iteration.executionReport = execResult.report;

      if (!execResult.report.success) {
        const errors = execResult.report.errors.map((e) => `${e.operation}: ${e.error}`);
        logger.warn('Patch execution had errors', { errors });
        messages = appendIterationContext(messages, i + 1, JSON.stringify(execResult.report), errors);
        iterations.push(iteration);
        continue;
      }

      currentBuffer = execResult.buffer;

      const validation = fileType === 'docx'
        ? await validateDocx(currentBuffer)
        : await validatePptx(currentBuffer);

      iteration.validationReport = validation;

      if (!validation.isValid) {
        logger.warn('Validation failed', { errors: validation.errors });
        messages = appendIterationContext(messages, i + 1, JSON.stringify(execResult.report), validation.errors);
        iterations.push(iteration);
        continue;
      }

      logger.info('Patch applied and validated successfully');
      iterations.push(iteration);
      break;

    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      logger.error(`Iteration ${i + 1} failed`, { error: errorMsg });
      iteration.error = errorMsg;
      messages = appendIterationContext(messages, i + 1, undefined, [errorMsg]);
      iterations.push(iteration);
    }
  }

  const lastIteration = iterations[iterations.length - 1];
  const success = lastIteration?.validationReport?.isValid ?? false;

  if (totalUsage.cacheHitTokens !== undefined && totalUsage.cacheMissTokens !== undefined) {
    const total = totalUsage.cacheHitTokens + totalUsage.cacheMissTokens;
    totalUsage.cacheHitRatio = total > 0 ? totalUsage.cacheHitTokens / total : 0;
  }
  totalUsage.stablePrefixHash = promptResult.stablePrefixHash;

  return {
    success,
    modifiedFile: success ? currentBuffer : undefined,
    manifest,
    iterations,
    summary: success
      ? `Successfully modified document after ${iterations.length} iteration(s).`
      : `Failed to modify document after ${iterations.length} iteration(s).`,
    usage: totalUsage,
  };
}
