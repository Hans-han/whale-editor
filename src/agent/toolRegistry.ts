import type { ToolDefinition } from '../types/index.js';

export const APPLY_PATCH_TOOL: ToolDefinition = {
  type: 'function',
  function: {
    name: 'apply_patch',
    description: 'Apply a structured patch to modify the Office document',
    parameters: {
      type: 'object',
      properties: {
        patch: {
          type: 'object',
          description: 'The patch envelope containing operations to apply',
          properties: {
            fileType: {
              type: 'string',
              enum: ['docx', 'pptx'],
              description: 'Type of Office file',
            },
            intent: {
              type: 'string',
              enum: ['content_edit', 'style_edit', 'layout_edit', 'structural_edit', 'validation_repair'],
              description: 'Intent of the patch',
            },
            operations: {
              type: 'array',
              description: 'Array of patch operations (max 20)',
              items: {
                type: 'object',
                properties: {
                  op: { type: 'string', description: 'Operation name' },
                  targetId: { type: 'string', description: 'Target object ID from manifest' },
                  payload: { type: 'object', description: 'Operation-specific payload' },
                  preserveFormatting: { type: 'boolean', description: 'Whether to preserve existing formatting' },
                },
                required: ['op', 'targetId', 'payload', 'preserveFormatting'],
              },
            },
            preserve: {
              type: 'object',
              properties: {
                styles: { type: 'boolean' },
                themes: { type: 'boolean' },
                relationships: { type: 'boolean' },
                comments: { type: 'boolean' },
                trackedChanges: { type: 'boolean' },
                speakerNotes: { type: 'boolean' },
              },
              required: ['styles', 'themes', 'relationships', 'comments', 'trackedChanges', 'speakerNotes'],
            },
          },
          required: ['fileType', 'intent', 'operations', 'preserve'],
        },
      },
      required: ['patch'],
    },
  },
};

export const AVAILABLE_TOOLS: ToolDefinition[] = [APPLY_PATCH_TOOL];
