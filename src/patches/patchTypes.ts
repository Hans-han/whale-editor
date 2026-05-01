import { z } from 'zod';

const PreserveOptionsSchema = z.object({
  styles: z.boolean(),
  themes: z.boolean(),
  relationships: z.boolean(),
  comments: z.boolean(),
  trackedChanges: z.boolean(),
  speakerNotes: z.boolean(),
});

const ReplaceParagraphTextSchema = z.object({
  op: z.literal('replace_paragraph_text'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ReplaceTextInParagraphSchema = z.object({
  op: z.literal('replace_text_in_paragraph'),
  targetId: z.string(),
  payload: z.object({ find: z.string(), replace: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const UpdateTableCellTextSchema = z.object({
  op: z.literal('update_table_cell_text'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const InsertParagraphAfterSchema = z.object({
  op: z.literal('insert_paragraph_after'),
  targetId: z.string(),
  payload: z.object({ text: z.string(), styleId: z.string().optional() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const DeleteParagraphSchema = z.object({
  op: z.literal('delete_paragraph'),
  targetId: z.string(),
  payload: z.object({}),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ApplyParagraphStyleSchema = z.object({
  op: z.literal('apply_paragraph_style'),
  targetId: z.string(),
  payload: z.object({ styleId: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const UpdateHeaderTextSchema = z.object({
  op: z.literal('update_header_text'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const UpdateFooterTextSchema = z.object({
  op: z.literal('update_footer_text'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ReplaceShapeTextSchema = z.object({
  op: z.literal('replace_shape_text'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ReplaceSlideTitleSchema = z.object({
  op: z.literal('replace_slide_title'),
  targetId: z.string(),
  payload: z.object({ title: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const UpdateSpeakerNotesSchema = z.object({
  op: z.literal('update_speaker_notes'),
  targetId: z.string(),
  payload: z.object({ text: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const InsertSlideFromLayoutSchema = z.object({
  op: z.literal('insert_slide_from_layout'),
  targetId: z.string(),
  payload: z.object({ layoutId: z.string(), position: z.number() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const DeleteSlideSchema = z.object({
  op: z.literal('delete_slide'),
  targetId: z.string(),
  payload: z.object({}),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const MoveShapeSchema = z.object({
  op: z.literal('move_shape'),
  targetId: z.string(),
  payload: z.object({ x: z.number(), y: z.number() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ResizeShapeSchema = z.object({
  op: z.literal('resize_shape'),
  targetId: z.string(),
  payload: z.object({ width: z.number(), height: z.number() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ApplyTextStyleSchema = z.object({
  op: z.literal('apply_text_style'),
  targetId: z.string(),
  payload: z.object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    fontSize: z.number().optional(),
    fontColor: z.string().optional(),
  }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const ReplaceImageSchema = z.object({
  op: z.literal('replace_image'),
  targetId: z.string(),
  payload: z.object({ imageData: z.string(), mimeType: z.string() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

const FitTextToShapeSchema = z.object({
  op: z.literal('fit_text_to_shape'),
  targetId: z.string(),
  payload: z.object({ text: z.string(), autoFit: z.boolean() }),
  preserveFormatting: z.boolean(),
  validationExpectations: z.array(z.string()).optional(),
});

export const PatchOperationSchema = z.discriminatedUnion('op', [
  ReplaceParagraphTextSchema,
  ReplaceTextInParagraphSchema,
  UpdateTableCellTextSchema,
  InsertParagraphAfterSchema,
  DeleteParagraphSchema,
  ApplyParagraphStyleSchema,
  UpdateHeaderTextSchema,
  UpdateFooterTextSchema,
  ReplaceShapeTextSchema,
  ReplaceSlideTitleSchema,
  UpdateSpeakerNotesSchema,
  InsertSlideFromLayoutSchema,
  DeleteSlideSchema,
  MoveShapeSchema,
  ResizeShapeSchema,
  ApplyTextStyleSchema,
  ReplaceImageSchema,
  FitTextToShapeSchema,
]);

export const PatchEnvelopeSchema = z.object({
  fileType: z.enum(['docx', 'pptx']),
  intent: z.enum(['content_edit', 'style_edit', 'layout_edit', 'structural_edit', 'validation_repair']),
  operations: z.array(PatchOperationSchema).max(20),
  preserve: PreserveOptionsSchema,
});

export type ValidatedPatchEnvelope = z.infer<typeof PatchEnvelopeSchema>;
export type ValidatedPatchOperation = z.infer<typeof PatchOperationSchema>;
