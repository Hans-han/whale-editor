export type FileType = 'docx' | 'pptx';

export interface DocxManifest {
  fileType: 'docx';
  documentId: string;
  sections: SectionInfo[];
  paragraphs: ParagraphInfo[];
  tables: TableInfo[];
  headers: HeaderFooterInfo[];
  footers: HeaderFooterInfo[];
  comments?: CommentInfo[];
  hasComments: boolean;
  hasFootnotes: boolean;
  hasEndnotes: boolean;
  hasNumbering: boolean;
  relationships: RelationshipSummary;
  media: MediaSummary;
  validation: PackageValidation;
}

export interface PptxManifest {
  fileType: 'pptx';
  deckId: string;
  slides: SlideInfo[];
  layouts: LayoutInfo[];
  masters: MasterInfo[];
  themes: ThemeInfo[];
  hasSpeakerNotes: boolean;
  hasComments: boolean;
  relationships: RelationshipSummary;
  media: MediaSummary;
  validation: PackageValidation;
}

export type OfficeManifest = DocxManifest | PptxManifest;

export interface SectionInfo {
  sectionId: string;
  startIndex: number;
  endIndex: number;
}

export interface ParagraphInfo {
  paragraphId: string;
  textPreview: string;
  styleId?: string;
  styleName?: string;
  headingLevel?: number;
  isListItem: boolean;
  listId?: string;
  numbering?: NumberingInfo;
  format?: ParagraphFormatInfo;
  runs?: RunInfo[];
  runCount?: number;
  commentIds?: string[];
}

export interface TableInfo {
  tableId: string;
  geometry?: TableGeometryInfo;
  rows: TableRowInfo[];
}

export interface TableRowInfo {
  rowId: string;
  cells: TableCellInfo[];
}

export interface TableCellInfo {
  cellId: string;
  textPreview: string;
  columnSpan: number;
  rowSpan: number;
  width?: MeasurementInfo;
  verticalMerge?: 'restart' | 'continue';
}

export interface HeaderFooterInfo {
  id: string;
  type: 'header' | 'footer';
  textPreview: string;
}

export interface CommentInfo {
  commentId: string;
  author?: string;
  date?: string;
  textPreview: string;
  anchoredParagraphIds: string[];
}

export interface RunInfo {
  runId: string;
  textPreview: string;
  styleId?: string;
  font?: string;
  fontSizeHalfPoints?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

export interface NumberingInfo {
  numId?: string;
  level?: string;
}

export interface ParagraphFormatInfo {
  alignment?: string;
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: string;
  };
  indentation?: {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
  };
}

export interface MeasurementInfo {
  type?: string;
  value?: number;
}

export interface TableGeometryInfo {
  width?: MeasurementInfo;
  gridColumns?: number[];
}

export interface SlideInfo {
  slideId: string;
  slideIndex: number;
  titleCandidate?: string;
  layoutId?: string;
  masterId?: string;
  shapes: ShapeInfo[];
  hasNotes: boolean;
  hasComments: boolean;
}

export interface ShapeInfo {
  shapeId: string;
  shapeType: string;
  placeholderType?: string;
  textPreview: string;
  position: ShapePosition;
}

export interface ShapePosition {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface LayoutInfo {
  layoutId: string;
  name: string;
  slideMasterId: string;
}

export interface MasterInfo {
  masterId: string;
  name: string;
  layoutIds: string[];
}

export interface ThemeInfo {
  themeId: string;
  name: string;
}

export interface RelationshipSummary {
  totalCount: number;
  types: Record<string, number>;
  externalLinks: string[];
}

export interface MediaSummary {
  totalCount: number;
  images: number;
  audio: number;
  video: number;
  totalSizeBytes: number;
}

export interface PackageValidation {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

export interface PatchEnvelope {
  fileType: FileType;
  intent: PatchIntent;
  operations: PatchOperation[];
  preserve: PreserveOptions;
}

export type PatchIntent =
  | 'content_edit'
  | 'style_edit'
  | 'layout_edit'
  | 'structural_edit'
  | 'validation_repair';

export interface PreserveOptions {
  styles: boolean;
  themes: boolean;
  relationships: boolean;
  comments: boolean;
  trackedChanges: boolean;
  speakerNotes: boolean;
}

export type PatchOperation =
  | ReplaceParagraphTextOp
  | ReplaceTextInParagraphOp
  | UpdateTableCellTextOp
  | InsertParagraphAfterOp
  | DeleteParagraphOp
  | ApplyParagraphStyleOp
  | UpdateHeaderTextOp
  | UpdateFooterTextOp
  | ReplaceShapeTextOp
  | ReplaceSlideTitleOp
  | UpdateSpeakerNotesOp
  | InsertSlideFromLayoutOp
  | DeleteSlideOp
  | MoveShapeOp
  | ResizeShapeOp
  | ApplyTextStyleOp
  | ReplaceImageOp
  | FitTextToShapeOp;

export interface ReplaceParagraphTextOp {
  op: 'replace_paragraph_text';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ReplaceTextInParagraphOp {
  op: 'replace_text_in_paragraph';
  targetId: string;
  payload: { find: string; replace: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface UpdateTableCellTextOp {
  op: 'update_table_cell_text';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface InsertParagraphAfterOp {
  op: 'insert_paragraph_after';
  targetId: string;
  payload: { text: string; styleId?: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface DeleteParagraphOp {
  op: 'delete_paragraph';
  targetId: string;
  payload: {};
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ApplyParagraphStyleOp {
  op: 'apply_paragraph_style';
  targetId: string;
  payload: { styleId: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface UpdateHeaderTextOp {
  op: 'update_header_text';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface UpdateFooterTextOp {
  op: 'update_footer_text';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ReplaceShapeTextOp {
  op: 'replace_shape_text';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ReplaceSlideTitleOp {
  op: 'replace_slide_title';
  targetId: string;
  payload: { title: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface UpdateSpeakerNotesOp {
  op: 'update_speaker_notes';
  targetId: string;
  payload: { text: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface InsertSlideFromLayoutOp {
  op: 'insert_slide_from_layout';
  targetId: string;
  payload: { layoutId: string; position: number };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface DeleteSlideOp {
  op: 'delete_slide';
  targetId: string;
  payload: {};
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface MoveShapeOp {
  op: 'move_shape';
  targetId: string;
  payload: { x: number; y: number };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ResizeShapeOp {
  op: 'resize_shape';
  targetId: string;
  payload: { width: number; height: number };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ApplyTextStyleOp {
  op: 'apply_text_style';
  targetId: string;
  payload: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: number;
    fontColor?: string;
  };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ReplaceImageOp {
  op: 'replace_image';
  targetId: string;
  payload: { imageData: string; mimeType: string };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface FitTextToShapeOp {
  op: 'fit_text_to_shape';
  targetId: string;
  payload: { text: string; autoFit: boolean };
  preserveFormatting: boolean;
  validationExpectations?: string[];
}

export interface ExecutionReport {
  success: boolean;
  operationsExecuted: number;
  operationsFailed: number;
  errors: ExecutionError[];
  warnings: string[];
}

export interface ExecutionError {
  operationIndex: number;
  operation: string;
  targetId: string;
  error: string;
  recoverable: boolean;
}

export interface AgentConfig {
  maxIterations: number;
  maxPatchOperations: number;
  maxOutputTokens: number;
  model: string;
  baseUrl: string;
  apiKey: string;
}

export interface AgentResult {
  success: boolean;
  modifiedFile?: Buffer;
  manifest: OfficeManifest;
  iterations: IterationLog[];
  summary: string;
  usage: UsageMetrics;
}

export interface IterationLog {
  iteration: number;
  patchPlan?: PatchEnvelope;
  executionReport?: ExecutionReport;
  validationReport?: PackageValidation;
  error?: string;
}

export interface LLMMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string;
  toolCallId?: string;
  toolCalls?: ToolCall[];
}

export interface ToolCall {
  id: string;
  type: 'function';
  function: {
    name: string;
    arguments: string;
  };
}

export interface ToolDefinition {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface GenerateOptions {
  temperature?: number;
  maxTokens?: number;
  toolChoice?: 'auto' | 'none' | 'required' | { type: 'function'; function: { name: string } };
}

export interface GenerateResult {
  content: string | null;
  toolCalls?: ToolCall[];
  usage: UsageMetrics;
}

export interface UsageMetrics {
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheHitTokens?: number;
  cacheMissTokens?: number;
  cacheHitRatio?: number;
  stablePrefixHash?: string;
}

export interface CacheMetrics {
  stablePrefixHash: string;
  cacheHitTokens: number;
  cacheMissTokens: number;
  cacheHitRatio: number;
}

export interface LLMProvider {
  name: string;
  generateText(messages: LLMMessage[], options?: GenerateOptions): Promise<GenerateResult>;
  generateToolCall(
    messages: LLMMessage[],
    tools: ToolDefinition[],
    options?: GenerateOptions
  ): Promise<GenerateResult>;
}

export interface PromptConfig {
  systemRules: string;
  toolSpecs: string;
  patchSchema: string;
  validationRules: string;
  manifest: OfficeManifest;
  userTask: string;
}

export interface PromptResult {
  messages: LLMMessage[];
  stablePrefixHash: string;
  cacheMetrics: CacheMetrics;
}
