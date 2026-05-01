import { loadPackage, parseXml, buildXml } from '../ooxmlPackage.js';
import type {
  PatchEnvelope,
  PatchOperation,
  ExecutionReport,
  ExecutionError,
} from '../../types/index.js';

interface XmlNode {
  [key: string]: unknown;
}

function toArray<T>(val: T | T[] | undefined): T[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

function parseTargetId(targetId: string): { slideNum: number; shapeNum?: number } {
  const parts = targetId.split('.');
  let slideNum = 0;
  let shapeNum: number | undefined;

  for (let i = 0; i < parts.length; i++) {
    if (parts[i] === 'slide' && i + 1 < parts.length) {
      slideNum = parseInt(parts[i + 1], 10);
    }
    if (parts[i] === 'shape' && i + 1 < parts.length) {
      shapeNum = parseInt(parts[i + 1], 10);
    }
  }

  return { slideNum, shapeNum };
}

function findShapeByIndex(spTree: XmlNode, index: number): XmlNode | null {
  const allShapes: XmlNode[] = [
    ...toArray(spTree['p:sp'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:pic'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:grpSp'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:cxnSp'] as XmlNode | XmlNode[]),
  ];

  if (index < 1 || index > allShapes.length) return null;
  return allShapes[index - 1];
}

function replaceShapeTextContent(shape: XmlNode, text: string): void {
  const txBody = shape['p:txBody'] ?? shape['a:txBody'];
  if (!txBody) {
    shape['p:txBody'] = {
      'a:p': {
        'a:r': {
          'a:t': text,
        },
      },
    };
    return;
  }

  const paras = toArray((txBody as XmlNode)['a:p'] as XmlNode | XmlNode[]);
  if (paras.length === 0) {
    (txBody as XmlNode)['a:p'] = {
      'a:r': { 'a:t': text },
    };
    return;
  }

  const firstPara = paras[0];
  const runs = toArray(firstPara['a:r'] as XmlNode | XmlNode[]);

  if (runs.length > 0) {
    if (!runs[0]['a:t']) runs[0]['a:t'] = {};
    runs[0]['a:t'] = text;
  } else {
    firstPara['a:r'] = { 'a:t': text };
  }
}

export async function applyPptxPatch(
  buffer: Buffer,
  patch: PatchEnvelope
): Promise<{ buffer: Buffer; report: ExecutionReport }> {
  if (patch.fileType !== 'pptx') {
    throw new Error('Invalid patch: fileType must be pptx');
  }

  const pkg = await loadPackage(buffer);
  const errors: ExecutionError[] = [];
  let executed = 0;
  let failed = 0;

  for (let i = 0; i < patch.operations.length; i++) {
    const op = patch.operations[i];
    try {
      await applyOperation(pkg, op);
      executed++;
    } catch (err) {
      failed++;
      errors.push({
        operationIndex: i,
        operation: op.op,
        targetId: op.targetId,
        error: err instanceof Error ? err.message : String(err),
        recoverable: false,
      });
    }
  }

  const newBuffer = await pkg.getBuffer();

  return {
    buffer: newBuffer,
    report: {
      success: failed === 0,
      operationsExecuted: executed,
      operationsFailed: failed,
      errors,
      warnings: [],
    },
  };
}

async function applyOperation(
  pkg: Awaited<ReturnType<typeof loadPackage>>,
  op: PatchOperation
): Promise<void> {
  const { slideNum, shapeNum } = parseTargetId(op.targetId);

  switch (op.op) {
    case 'replace_shape_text': {
      if (!slideNum || !shapeNum) throw new Error(`Invalid target: ${op.targetId}`);
      const slideXml = await pkg.getContent(`ppt/slides/slide${slideNum}.xml`);
      if (!slideXml) throw new Error(`Slide ${slideNum} not found`);

      const parsed = parseXml(slideXml) as XmlNode;
      const sld = parsed['p:sld'] as XmlNode;
      const cSld = sld['p:cSld'] as XmlNode;
      const spTree = cSld['p:spTree'] as XmlNode;

      const shape = findShapeByIndex(spTree, shapeNum);
      if (!shape) throw new Error(`Shape ${shapeNum} not found on slide ${slideNum}`);

      replaceShapeTextContent(shape, op.payload.text);
      await pkg.setContent(`ppt/slides/slide${slideNum}.xml`, buildXml(parsed));
      break;
    }

    case 'replace_slide_title': {
      if (!slideNum) throw new Error(`Invalid target: ${op.targetId}`);
      const slideXml = await pkg.getContent(`ppt/slides/slide${slideNum}.xml`);
      if (!slideXml) throw new Error(`Slide ${slideNum} not found`);

      const parsed = parseXml(slideXml) as XmlNode;
      const sld = parsed['p:sld'] as XmlNode;
      const cSld = sld['p:cSld'] as XmlNode;
      const spTree = cSld['p:spTree'] as XmlNode;

      const shapes = [
        ...toArray(spTree['p:sp'] as XmlNode | XmlNode[]),
      ];

      let titleShape: XmlNode | null = null;
      for (const shape of shapes) {
        const nvSpPr = shape['p:nvSpPr'] as XmlNode;
        const nvPr = nvSpPr?.['p:nvPr'] as XmlNode;
        const ph = nvPr?.['p:ph'] as XmlNode;
        if (ph?.['@_type'] === 'title' || ph?.['@_type'] === 'ctrTitle') {
          titleShape = shape;
          break;
        }
      }

      if (!titleShape && shapes.length > 0) {
        titleShape = shapes[0];
      }

      if (!titleShape) throw new Error('No title shape found');

      replaceShapeTextContent(titleShape, op.payload.title);
      await pkg.setContent(`ppt/slides/slide${slideNum}.xml`, buildXml(parsed));
      break;
    }

    case 'update_speaker_notes': {
      if (!slideNum) throw new Error(`Invalid target: ${op.targetId}`);
      const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;

      if (!pkg.hasPath(notesPath)) {
        throw new Error(`Notes for slide ${slideNum} not found`);
      }

      const notesXml = await pkg.getContent(notesPath);
      if (!notesXml) throw new Error(`Failed to read notes for slide ${slideNum}`);

      const parsed = parseXml(notesXml) as XmlNode;
      const notes = parsed['p:notes'] as XmlNode;
      const cSld = notes['p:cSld'] as XmlNode;
      const spTree = cSld['p:spTree'] as XmlNode;

      const shapes = toArray(spTree['p:sp'] as XmlNode | XmlNode[]);
      let notesShape: XmlNode | null = null;

      for (const shape of shapes) {
        const nvSpPr = shape['p:nvSpPr'] as XmlNode;
        const nvPr = nvSpPr?.['p:nvPr'] as XmlNode;
        const ph = nvPr?.['p:ph'] as XmlNode;
        if (ph?.['@_type'] === 'body') {
          notesShape = shape;
          break;
        }
      }

      if (!notesShape && shapes.length > 0) {
        notesShape = shapes[shapes.length - 1];
      }

      if (!notesShape) throw new Error('No notes body shape found');

      replaceShapeTextContent(notesShape, op.payload.text);
      await pkg.setContent(notesPath, buildXml(parsed));
      break;
    }

    case 'delete_slide': {
      if (!slideNum) throw new Error(`Invalid target: ${op.targetId}`);

      const presXml = await pkg.getContent('ppt/presentation.xml');
      if (!presXml) throw new Error('presentation.xml not found');

      const parsed = parseXml(presXml) as XmlNode;
      const presentation = parsed['p:presentation'] as XmlNode;
      const sldIdLst = presentation['p:sldIdLst'] as XmlNode;
      const sldIds = toArray(sldIdLst['p:sldId'] as XmlNode | XmlNode[]);

      if (slideNum < 1 || slideNum > sldIds.length) {
        throw new Error(`Slide ${slideNum} not found`);
      }

      sldIds.splice(slideNum - 1, 1);
      sldIdLst['p:sldId'] = sldIds;

      await pkg.setContent('ppt/presentation.xml', buildXml(parsed));
      break;
    }

    default:
      throw new Error(`Unsupported operation: ${op.op}`);
  }
}
