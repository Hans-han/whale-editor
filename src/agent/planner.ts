import type {
  LLMProvider,
  OfficeManifest,
  ToolDefinition,
  FileType,
  UsageMetrics,
} from '../types/index.js';
import { stableStringify } from '../utils/stableJson.js';
import { DeepSeekProvider } from '../llm/deepseekProvider.js';
import { logger } from '../utils/logger.js';

// Planner-specific model. Defaults to deepseek-chat (non-thinking) because:
// - planner is a structured tool-call task, doesn't benefit from reasoning
// - thinking models on DeepSeek can stall or refuse tool_choice variants
// - non-thinking is faster and more reliable for this stage
// Override via DEEPSEEK_PLANNER_MODEL env if needed.
const PLANNER_MODEL = process.env.DEEPSEEK_PLANNER_MODEL ?? 'deepseek-chat';

export interface PlanStep {
  title: string;
  task: string;
}

export interface AgentPlan {
  rationale: string;
  steps: PlanStep[];
}

const PLAN_TOOL: ToolDefinition = {
  type: 'function',
  function: {
    name: 'submit_plan',
    description:
      'Submit a step-by-step execution plan for editing the Office document.',
    parameters: {
      type: 'object',
      properties: {
        rationale: {
          type: 'string',
          description: 'Brief high-level reasoning for the plan (1-3 sentences).',
        },
        steps: {
          type: 'array',
          description: 'Ordered list of atomic, executable steps.',
          items: {
            type: 'object',
            properties: {
              title: {
                type: 'string',
                description: 'Short user-facing title (≤30 chars).',
              },
              task: {
                type: 'string',
                description:
                  'Concrete instruction for the executor agent. Reference manifest IDs when possible.',
              },
            },
            required: ['title', 'task'],
          },
        },
      },
      required: ['rationale', 'steps'],
    },
  },
};

const PLANNER_SYSTEM = `You are a senior planning agent for Office-native document editing.

Given a document manifest (DOCX or PPTX) and the user's intent (a free-form instruction, format spec, or template-following request), produce a step-by-step execution plan that an executor agent can carry out via structured patch operations.

Language:
- Return rationale, every step title, and every step task in Simplified Chinese.
- Keep manifest IDs, operation names, style IDs, exact quoted document text, and file/object identifiers unchanged.
- Do not write English prose in user-facing fields unless the English words are exact document content such as "Abstract" or "Keywords:".

Executor capabilities (DOCX): replace_paragraph_text, replace_text_in_paragraph, update_table_cell_text, insert_paragraph_after, delete_paragraph, apply_paragraph_style, update_header_text, update_footer_text.
Executor capabilities (PPTX): replace_shape_text, replace_slide_title, update_speaker_notes, insert_slide_from_layout, delete_slide, move_shape, resize_shape, apply_text_style, replace_image, fit_text_to_shape.

Input matching:
- The user intent may include a fixed input-matching contract plus multiple input blocks from typed text, PDFs, reference documents, templates, or extracted comments.
- Apply that contract before planning. Classify sources as DIRECT_EDIT, STYLE_SPEC, TEMPLATE_SOURCE, CONTENT_SOURCE, REVIEW_COMMENT, TRACKED_CHANGE_REVIEW, GLOBAL_CONSISTENCY, LOCAL_SELECTION, or AMBIGUOUS_CONTEXT.
- If the manifest includes comments, treat comment text and anchoredParagraphIds as review context. When the user asks to handle comments/批注, make one or more local steps tied to those anchors.
- If references are style guides or templates, plan structural/style matching without blindly copying all reference content.

Format fidelity standard:
- Prefer selection-like, minimal target ranges. Do not plan whole-document rewrites unless the user explicitly asks for a global transformation.
- Use manifest IDs and local style context. Preserve styleId, styleName, headingLevel, listId, table geometry, slide layout, master, theme, notes, comments, and relationships unless explicitly asked to change them.
- Use low-level DOCX fields when present: numbering.numId, numbering.level, format.spacing, format.indentation, runs[].font, runs[].fontSizeHalfPoints, runs[].bold, runs[].italic, runs[].color, table geometry widths, and cell widths.
- For DOCX text edits, plan substring replacements before paragraph rewrites when possible.
- If a DOCX target phrase appears to span multiple runs with different styles, keep the step narrow and ask the executor to note the mixed-run formatting risk in validationExpectations.
- For DOCX table edits, plan cell updates in place and avoid structural table changes.
- For DOCX insertions, reference the paragraph whose style should be inherited.
- For PPTX, keep shape positions and sizes unless the user explicitly asks for layout changes. Prefer fit_text_to_shape if overflow is likely.
- For template-following requests, fill existing placeholders and analogous sections before inserting new structure.
- Do not claim native tracked-change creation; the executor preserves review context and returns auditable patch operations.

Rules:
1. Each step is atomic and self-contained — an executor receives only the step's task string plus the document, no other context.
2. Reference IDs from the manifest where possible (paragraphId, tableCellId, slideId, shapeId).
3. Step titles ≤30 chars, shown to the user as progress labels, and must be Chinese.
4. Step tasks are clear, concrete Chinese instructions and include preservation constraints such as "保留 styleId/listId/table geometry" when relevant.
5. Order steps so earlier ones don't conflict with later ones.
6. Keep plans focused: 1–8 steps typically; never more than 12.
7. Do not invent content not implied by the user intent.
8. If the intent is ambiguous, pick the most reasonable concrete interpretation rather than asking back.

Output ONLY via the submit_plan tool.`;

/**
 * Try to recover a plan JSON from a free-form assistant message.
 * Handles three common reasoner output shapes:
 *   1. Pure JSON: `{ "rationale": ..., "steps": [...] }`
 *   2. Fenced: ```json\n{...}\n```
 *   3. Embedded inside prose
 */
function extractPlanFromText(content: string): { rationale?: string; steps?: PlanStep[] } | null {
  if (!content) return null;
  const tryParse = (s: string): unknown => {
    try {
      return JSON.parse(s);
    } catch {
      return null;
    }
  };
  const candidates: string[] = [];

  // Fenced code blocks
  const fenceRe = /```(?:json)?\s*([\s\S]*?)```/gi;
  for (const m of content.matchAll(fenceRe)) {
    if (m[1]) candidates.push(m[1].trim());
  }
  // Whole content (in case it's pure JSON)
  candidates.push(content.trim());
  // First {...} block, naive
  const braceStart = content.indexOf('{');
  const braceEnd = content.lastIndexOf('}');
  if (braceStart !== -1 && braceEnd > braceStart) {
    candidates.push(content.slice(braceStart, braceEnd + 1));
  }

  for (const c of candidates) {
    const parsed = tryParse(c);
    if (parsed && typeof parsed === 'object' && Array.isArray((parsed as any).steps)) {
      return parsed as { rationale?: string; steps?: PlanStep[] };
    }
  }
  return null;
}

export async function planTasks(
  provider: LLMProvider,
  manifest: OfficeManifest,
  intent: string,
  fileType: FileType
): Promise<{ plan: AgentPlan; usage: UsageMetrics }> {
  const manifestJson = stableStringify(manifest);

  const messages = [
    { role: 'system' as const, content: PLANNER_SYSTEM },
    {
      role: 'user' as const,
      content: [
        '## Document Manifest\n\n```json\n',
        manifestJson,
        '\n```\n\n',
        `## User Intent (${fileType})\n\n`,
        intent,
        '\n\nProduce an execution plan with 1–8 atomic steps. Output via submit_plan only.',
      ].join(''),
    },
  ];

  // Use a dedicated planner provider with a non-thinking model
  // (deepseek-chat by default). Inherit the user/server API key from the
  // calling provider so BYO-key still works.
  const plannerProvider =
    provider instanceof DeepSeekProvider
      ? new DeepSeekProvider({
          apiKey: (provider as unknown as { config: { apiKey: string } }).config.apiKey,
          baseUrl: (provider as unknown as { config: { baseUrl: string } }).config.baseUrl,
          model: PLANNER_MODEL,
        })
      : provider;

  logger.info('Calling planner LLM', {
    fileType,
    intentLength: intent.length,
    plannerModel: PLANNER_MODEL,
  });

  const result = await plannerProvider.generateToolCall(messages, [PLAN_TOOL], {
    temperature: 0.2,
    maxTokens: 4096,
    toolChoice: 'auto',
  });

  let args: { rationale?: string; steps?: PlanStep[] } | null = null;

  if (result.toolCalls && result.toolCalls.length > 0) {
    const toolCall = result.toolCalls[0];
    if (toolCall.function.name !== 'submit_plan') {
      throw new Error(`Unexpected tool call: ${toolCall.function.name}`);
    }
    try {
      args = JSON.parse(toolCall.function.arguments);
    } catch (err) {
      logger.warn('Planner tool args parse failed, will try content fallback', {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  // Fallback: reasoner/thinking models sometimes emit the JSON plan as
  // assistant content instead of a tool call. Try to recover.
  if (!args && result.content) {
    const fallback = extractPlanFromText(result.content);
    if (fallback) {
      logger.info('Planner: recovered plan from content (no tool call)');
      args = fallback;
    }
  }

  if (!args) {
    logger.error('Planner failed to produce plan', {
      hasToolCall: !!result.toolCalls?.length,
      contentPreview: result.content?.slice(0, 200) ?? '(no content)',
    });
    throw new Error(
      'Planner did not return a usable plan (neither tool call nor parseable JSON content). ' +
        'Try rephrasing the intent more concretely.'
    );
  }

  const plan: AgentPlan = {
    rationale: typeof args.rationale === 'string' ? args.rationale : '',
    steps: Array.isArray(args.steps)
      ? args.steps
          .filter(
            (s): s is PlanStep =>
              !!s &&
              typeof s.title === 'string' &&
              typeof s.task === 'string' &&
              s.task.trim().length > 0
          )
          .map((s) => ({ title: s.title.slice(0, 60), task: s.task }))
      : [],
  };

  if (plan.steps.length === 0) {
    throw new Error('Planner returned empty plan');
  }
  if (plan.steps.length > 12) {
    logger.warn(`Planner returned ${plan.steps.length} steps, truncating to 12`);
    plan.steps = plan.steps.slice(0, 12);
  }

  logger.info('Planner produced plan', { steps: plan.steps.length });

  return { plan, usage: result.usage };
}
