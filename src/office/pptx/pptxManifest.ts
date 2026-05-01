import { loadPackage, parseXml } from '../ooxmlPackage.js';
import type {
  PptxManifest,
  SlideInfo,
  ShapeInfo,
  ShapePosition,
  LayoutInfo,
  MasterInfo,
  ThemeInfo,
  RelationshipSummary,
  MediaSummary,
} from '../../types/index.js';

interface XmlNode {
  [key: string]: unknown;
}

function toArray<T>(val: T | T[] | undefined): T[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

function extractText(node: unknown): string {
  if (!node) return '';
  if (typeof node === 'string') return node;
  if (typeof node !== 'object') return String(node);

  const obj = node as Record<string, unknown>;
  if ('a:t' in obj) {
    const t = obj['a:t'];
    if (typeof t === 'string') return t;
    if (t && typeof t === 'object' && '#text' in (t as Record<string, unknown>)) {
      return String((t as Record<string, unknown>)['#text']);
    }
  }

  let result = '';
  for (const key of Object.keys(obj)) {
    if (key.startsWith('@_')) continue;
    const val = obj[key];
    if (Array.isArray(val)) {
      for (const item of val) {
        result += extractText(item);
      }
    } else if (val && typeof val === 'object') {
      result += extractText(val);
    }
  }
  return result;
}

function generateId(prefix: string, index: number): string {
  return `${prefix}.${String(index).padStart(4, '0')}`;
}

function extractShapePosition(shape: XmlNode): ShapePosition {
  const spPr = shape['p:spPr'] as XmlNode | undefined;
  const xfrm = spPr?.['a:xfrm'] as XmlNode | undefined;
  const off = xfrm?.['a:off'] as XmlNode | undefined;
  const ext = xfrm?.['a:ext'] as XmlNode | undefined;

  return {
    x: Number(off?.['@_x'] ?? 0),
    y: Number(off?.['@_y'] ?? 0),
    width: Number(ext?.['@_cx'] ?? 0),
    height: Number(ext?.['@_cy'] ?? 0),
  };
}

function getShapeType(shape: XmlNode): string {
  if (shape['p:sp']) return 'shape';
  if (shape['p:pic']) return 'picture';
  if (shape['p:grpSp']) return 'group';
  if (shape['p:cxnSp']) return 'connector';
  return 'unknown';
}

function getPlaceholderType(shape: XmlNode): string | undefined {
  const nvSpPr = shape['p:nvSpPr'] as XmlNode | undefined;
  const nvPr = nvSpPr?.['p:nvPr'] as XmlNode | undefined;
  const ph = nvPr?.['p:ph'] as XmlNode | undefined;
  return ph?.['@_type'] as string | undefined;
}

export async function generatePptxManifest(buffer: Buffer): Promise<PptxManifest> {
  const pkg = await loadPackage(buffer);
  const paths = pkg.listPaths();

  const presentationXml = await pkg.getContent('ppt/presentation.xml');
  if (!presentationXml) {
    throw new Error('Invalid PPTX: missing ppt/presentation.xml');
  }

  const parsed = parseXml(presentationXml) as XmlNode;
  const presentation = parsed['p:presentation'] as XmlNode;
  if (!presentation) {
    throw new Error('Invalid PPTX: missing p:presentation root');
  }

  const sldIdLst = presentation['p:sldIdLst'] as XmlNode | undefined;
  const sldIds = toArray(sldIdLst?.['p:sldId'] as XmlNode | XmlNode[]);

  const slides: SlideInfo[] = [];
  for (let i = 0; i < sldIds.length; i++) {
    const sldId = sldIds[i];
    const slideNum = i + 1;
    const slidePath = `ppt/slides/slide${slideNum}.xml`;

    const slideXml = await pkg.getContent(slidePath);
    if (!slideXml) continue;

    const slideParsed = parseXml(slideXml) as XmlNode;
    const sld = slideParsed['p:sld'] as XmlNode;
    if (!sld) continue;

    const cSld = sld['p:cSld'] as XmlNode;
    const spTree = cSld?.['p:spTree'] as XmlNode;
    const shapes = extractShapes(spTree);

    const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    const hasNotes = pkg.hasPath(notesPath);

    let titleCandidate: string | undefined;
    for (const shape of shapes) {
      if (shape.placeholderType === 'title' || shape.placeholderType === 'ctrTitle') {
        titleCandidate = shape.textPreview;
        break;
      }
    }
    if (!titleCandidate && shapes.length > 0) {
      titleCandidate = shapes[0].textPreview;
    }

    slides.push({
      slideId: generateId('pptx.slide', slideNum),
      slideIndex: slideNum,
      titleCandidate,
      shapes,
      hasNotes,
      hasComments: false,
    });
  }

  const layouts = await extractLayouts(pkg);
  const masters = await extractMasters(pkg);
  const themes = await extractThemes(pkg, paths);

  return {
    fileType: 'pptx',
    deckId: 'pptx.main',
    slides,
    layouts,
    masters,
    themes,
    hasSpeakerNotes: slides.some((s) => s.hasNotes),
    hasComments: false,
    relationships: extractRelationships(paths),
    media: extractMedia(paths),
    validation: { isValid: true, errors: [], warnings: [] },
  };
}

function extractShapes(spTree: XmlNode | undefined): ShapeInfo[] {
  if (!spTree) return [];

  const shapes: ShapeInfo[] = [];
  const children = [
    ...toArray(spTree['p:sp'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:pic'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:grpSp'] as XmlNode | XmlNode[]),
    ...toArray(spTree['p:cxnSp'] as XmlNode | XmlNode[]),
  ];

  for (let i = 0; i < children.length; i++) {
    const shape = children[i];
    const text = extractShapeText(shape);
    const position = extractShapePosition(shape);

    shapes.push({
      shapeId: `pptx.shape.${String(i + 1).padStart(4, '0')}`,
      shapeType: getShapeType(shape),
      placeholderType: getPlaceholderType(shape),
      textPreview: text.slice(0, 800),
      position,
    });
  }

  return shapes;
}

function extractShapeText(shape: XmlNode): string {
  const txBody = shape['p:txBody'] ?? shape['a:txBody'];
  if (!txBody) return '';
  return extractText(txBody);
}

async function extractLayouts(pkg: Awaited<ReturnType<typeof loadPackage>>): Promise<LayoutInfo[]> {
  const layouts: LayoutInfo[] = [];
  const paths = pkg.listPaths().filter((p) => p.startsWith('ppt/slideLayouts/slideLayout') && p.endsWith('.xml'));

  for (const path of paths) {
    const xml = await pkg.getContent(path);
    if (!xml) continue;

    const parsed = parseXml(xml) as XmlNode;
    const sldLayout = parsed['p:sldLayout'] as XmlNode;
    if (!sldLayout) continue;

    const num = parseInt(path.replace(/[^0-9]/g, ''), 10);
    layouts.push({
      layoutId: generateId('pptx.layout', num),
      name: sldLayout['@_type'] as string ?? 'unknown',
      slideMasterId: 'pptx.master.0001',
    });
  }

  return layouts;
}

async function extractMasters(pkg: Awaited<ReturnType<typeof loadPackage>>): Promise<MasterInfo[]> {
  const masters: MasterInfo[] = [];
  const paths = pkg.listPaths().filter((p) => p.startsWith('ppt/slideMasters/slideMaster') && p.endsWith('.xml'));

  for (const path of paths) {
    const xml = await pkg.getContent(path);
    if (!xml) continue;

    const num = parseInt(path.replace(/[^0-9]/g, ''), 10);
    masters.push({
      masterId: generateId('pptx.master', num),
      name: `Slide Master ${num}`,
      layoutIds: [],
    });
  }

  return masters;
}

async function extractThemes(pkg: Awaited<ReturnType<typeof loadPackage>>, paths: string[]): Promise<ThemeInfo[]> {
  const themePaths = paths.filter((p) => p.startsWith('ppt/theme/theme') && p.endsWith('.xml'));
  const themes: ThemeInfo[] = [];

  for (const path of themePaths) {
    const num = parseInt(path.replace(/[^0-9]/g, ''), 10);
    themes.push({
      themeId: generateId('pptx.theme', num),
      name: `Theme ${num}`,
    });
  }

  return themes;
}

function extractRelationships(paths: string[]): RelationshipSummary {
  const relPaths = paths.filter((p) => p.includes('_rels') && p.endsWith('.rels'));
  let totalCount = 0;
  const types: Record<string, number> = {};

  for (const path of relPaths) {
    totalCount++;
    const type = path.includes('presentation') ? 'presentation' : 'other';
    types[type] = (types[type] ?? 0) + 1;
  }

  return { totalCount, types, externalLinks: [] };
}

function extractMedia(paths: string[]): MediaSummary {
  const mediaPaths = paths.filter((p) => p.startsWith('ppt/media/'));
  let images = 0;
  let audio = 0;
  let video = 0;

  for (const path of mediaPaths) {
    const ext = path.split('.').pop()?.toLowerCase() ?? '';
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) {
      images++;
    } else if (['mp3', 'wav', 'ogg'].includes(ext)) {
      audio++;
    } else if (['mp4', 'avi', 'mov'].includes(ext)) {
      video++;
    }
  }

  return {
    totalCount: images + audio + video,
    images,
    audio,
    video,
    totalSizeBytes: 0,
  };
}
