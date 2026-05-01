import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { stableStringify } from '../utils/stableJson.js';
import { sha256 } from '../utils/hashing.js';
import type {
  LLMMessage,
  OfficeManifest,
  PromptConfig,
  PromptResult,
  CacheMetrics,
} from '../types/index.js';

const SYSTEM_PROMPT_PATH = resolve(import.meta.dirname, '../../prompts/office-agent-system.md');
let cachedSystemPrompt: string | null = null;

function getSystemPrompt(): string {
  if (cachedSystemPrompt) return cachedSystemPrompt;
  cachedSystemPrompt = readFileSync(SYSTEM_PROMPT_PATH, 'utf8');
  return cachedSystemPrompt;
}

export function buildPrompt(config: PromptConfig): PromptResult {
  const messages: LLMMessage[] = [];

  const systemContent = [
    config.systemRules || getSystemPrompt(),
    '\n\n## Available Tools\n\n',
    config.toolSpecs,
    '\n\n## Patch Schema\n\n',
    config.patchSchema,
    '\n\n## Validation Rules\n\n',
    config.validationRules,
  ].join('');

  messages.push({
    role: 'system',
    content: systemContent,
  });

  const manifestJson = stableStringify(config.manifest);

  const userContent = [
    '## Document Manifest\n\n```json\n',
    manifestJson,
    '\n```\n\n',
    '## User Task\n\n',
    config.userTask,
  ].join('');

  messages.push({
    role: 'user',
    content: userContent,
  });

  const stablePrefix = systemContent + userContent;
  const stablePrefixHash = sha256(stablePrefix);

  const cacheMetrics: CacheMetrics = {
    stablePrefixHash,
    cacheHitTokens: 0,
    cacheMissTokens: 0,
    cacheHitRatio: 0,
  };

  return {
    messages,
    stablePrefixHash,
    cacheMetrics,
  };
}

export function appendIterationContext(
  messages: LLMMessage[],
  iteration: number,
  toolResult?: string,
  validationErrors?: string[],
  changedIds?: string[]
): LLMMessage[] {
  const newMessages = [...messages];

  const dynamicContent = [
    `\n\n## Current Iteration: ${iteration}\n`,
  ];

  if (toolResult) {
    dynamicContent.push('\n\n### Previous Patch Result\n\n```\n', toolResult, '\n```');
  }

  if (validationErrors && validationErrors.length > 0) {
    dynamicContent.push('\n\n### Validation Errors\n\n');
    for (const error of validationErrors) {
      dynamicContent.push('- ', error, '\n');
    }
  }

  if (changedIds && changedIds.length > 0) {
    dynamicContent.push('\n\n### Changed Object IDs\n\n');
    for (const id of changedIds) {
      dynamicContent.push('- ', id, '\n');
    }
  }

  newMessages.push({
    role: 'user',
    content: dynamicContent.join(''),
  });

  return newMessages;
}

export const TOOL_SPECS = `
You have one tool available: \`apply_patch\`

This tool applies a structured patch to the Office document.

Parameters:
- patch: The complete patch envelope JSON (required)

Example usage:
\`\`\`json
{
  "patch": {
    "fileType": "docx",
    "intent": "content_edit",
    "operations": [...],
    "preserve": {...}
  }
}
\`\`\`
`;

export const PATCH_SCHEMA = `
Patch envelope must be a valid JSON object with:
- fileType: "docx" or "pptx"
- intent: one of "content_edit", "style_edit", "layout_edit", "structural_edit", "validation_repair"
- operations: array of operation objects (max 20)
- preserve: object with boolean flags for what to preserve
`;

export const VALIDATION_RULES = `
Validation checks:
1. ZIP package opens successfully
2. Required OOXML parts exist
3. XML parses without errors
4. Relationships targets exist
5. Modified target objects exist
6. Content types are valid
7. For PPTX: slide relationships intact
8. For DOCX: document.xml structure valid
`;
