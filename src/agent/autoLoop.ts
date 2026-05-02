import type {
  LLMProvider,
  FileType,
  UsageMetrics,
  OfficeManifest,
  PatchOperation,
} from '../types/index.js';
import { generateDocxManifest } from '../office/docx/docxManifest.js';
import { generatePptxManifest } from '../office/pptx/pptxManifest.js';
import { runAgent } from './loop.js';
import { planTasks, type AgentPlan, type PlanStep } from './planner.js';
import { logger } from '../utils/logger.js';

export interface SessionStatus {
  id: string;
  filename: string;
  fileType: FileType;
  reuseCount: number;
  reuseRemaining: number;
  expiresInMs: number;
  ttlMs: number;
}

export type AutoEvent =
  | { type: 'converting'; from: 'doc' | 'ppt'; filename: string }
  | { type: 'session'; session: SessionStatus }
  | { type: 'manifest_summary'; fileType: FileType; objectCount: number }
  | { type: 'plan'; plan: AgentPlan; usage: UsageMetrics }
  | { type: 'step_start'; index: number; total: number; step: PlanStep }
  | {
      type: 'step_done';
      index: number;
      success: boolean;
      iterations: number;
      summary: string;
      usage: UsageMetrics;
      operations?: PatchOperation[];
      intent?: string;
    }
  | { type: 'step_error'; index: number; error: string }
  | {
      type: 'complete';
      success: boolean;
      modifiedFile?: Buffer;
      modifiedFileBase64?: string;
      downloadUrl?: string;
      modifiedFilename?: string;
      totalUsage: UsageMetrics;
    }
  | { type: 'error'; error: string };

export type AutoEmit = (event: AutoEvent) => void;

function summarizeManifest(manifest: OfficeManifest): number {
  if (manifest.fileType === 'docx') {
    return (
      manifest.paragraphs.length +
      manifest.tables.reduce(
        (acc, t) => acc + t.rows.reduce((a, r) => a + r.cells.length, 0),
        0
      ) +
      manifest.headers.length +
      manifest.footers.length
    );
  }
  return manifest.slides.reduce((acc, s) => acc + s.shapes.length + 1, 0);
}

function accumulate(total: UsageMetrics, delta: UsageMetrics): void {
  total.inputTokens += delta.inputTokens;
  total.outputTokens += delta.outputTokens;
  total.totalTokens += delta.totalTokens;
  if (delta.cacheHitTokens !== undefined) {
    total.cacheHitTokens = (total.cacheHitTokens ?? 0) + delta.cacheHitTokens;
  }
  if (delta.cacheMissTokens !== undefined) {
    total.cacheMissTokens = (total.cacheMissTokens ?? 0) + delta.cacheMissTokens;
  }
}

function finalizeCacheRatio(usage: UsageMetrics): void {
  if (
    usage.cacheHitTokens !== undefined &&
    usage.cacheMissTokens !== undefined
  ) {
    const total = (usage.cacheHitTokens ?? 0) + (usage.cacheMissTokens ?? 0);
    usage.cacheHitRatio = total > 0 ? (usage.cacheHitTokens ?? 0) / total : 0;
  }
}

export async function runAuto(
  provider: LLMProvider,
  fileBuffer: Buffer,
  fileType: FileType,
  intent: string,
  emit: AutoEmit
): Promise<void> {
  const totalUsage: UsageMetrics = {
    inputTokens: 0,
    outputTokens: 0,
    totalTokens: 0,
    cacheHitTokens: 0,
    cacheMissTokens: 0,
  };

  try {
    logger.info('Auto run starting', { fileType, intentLength: intent.length });

    const manifest =
      fileType === 'docx'
        ? await generateDocxManifest(fileBuffer)
        : await generatePptxManifest(fileBuffer);

    emit({
      type: 'manifest_summary',
      fileType,
      objectCount: summarizeManifest(manifest),
    });

    const { plan, usage: planUsage } = await planTasks(
      provider,
      manifest,
      intent,
      fileType
    );
    accumulate(totalUsage, planUsage);
    emit({ type: 'plan', plan, usage: planUsage });

    let currentBuffer = fileBuffer;
    let allSuccess = true;

    for (let i = 0; i < plan.steps.length; i++) {
      const step = plan.steps[i];
      emit({ type: 'step_start', index: i, total: plan.steps.length, step });

      try {
        const result = await runAgent(
          provider,
          currentBuffer,
          fileType,
          step.task
        );
        accumulate(totalUsage, result.usage);

        // Pull the operations from the last successful iteration so the UI
        // can show the user what was actually changed.
        const lastIter = result.iterations[result.iterations.length - 1];
        const ops = lastIter?.patchPlan?.operations;
        const intent = lastIter?.patchPlan?.intent;

        if (result.success && result.modifiedFile) {
          currentBuffer = result.modifiedFile;
          emit({
            type: 'step_done',
            index: i,
            success: true,
            iterations: result.iterations.length,
            summary: result.summary,
            usage: result.usage,
            operations: ops,
            intent,
          });
        } else {
          allSuccess = false;
          emit({
            type: 'step_done',
            index: i,
            success: false,
            iterations: result.iterations.length,
            summary: result.summary,
            usage: result.usage,
            operations: ops,
            intent,
          });
        }
      } catch (err) {
        allSuccess = false;
        const msg = err instanceof Error ? err.message : String(err);
        logger.error(`Auto step ${i} threw`, { error: msg });
        emit({ type: 'step_error', index: i, error: msg });
      }
    }

    finalizeCacheRatio(totalUsage);

    emit({
      type: 'complete',
      success: allSuccess,
      modifiedFile: currentBuffer !== fileBuffer ? currentBuffer : undefined,
      totalUsage,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    logger.error('Auto run failed', { error: msg });
    finalizeCacheRatio(totalUsage);
    emit({ type: 'error', error: msg });
  }
}
