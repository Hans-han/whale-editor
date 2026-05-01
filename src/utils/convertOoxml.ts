import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, writeFile, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, basename } from 'node:path';
import { existsSync } from 'node:fs';
import { logger } from './logger.js';

const exec = promisify(execFile);

const SOFFICE_CANDIDATES = [
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
  '/usr/local/bin/soffice',
  '/opt/homebrew/bin/soffice',
  '/usr/bin/soffice',
];

let cachedBin: string | null | undefined = undefined;

export async function findSoffice(): Promise<string | null> {
  if (cachedBin !== undefined) return cachedBin;
  // PATH lookup
  try {
    const { stdout } = await exec('which', ['soffice']);
    const path = stdout.trim();
    if (path) {
      cachedBin = path;
      return path;
    }
  } catch {
    /* not in PATH */
  }
  for (const p of SOFFICE_CANDIDATES) {
    if (existsSync(p)) {
      cachedBin = p;
      return p;
    }
  }
  cachedBin = null;
  return null;
}

export interface ConvertedDoc {
  buffer: Buffer;
  filename: string; // .docx or .pptx
  kind: 'docx' | 'pptx';
}

const HELP_MSG =
  '检测不到 libreoffice，无法转换旧版 Office 二进制格式。\n请执行：brew install --cask libreoffice\n或在 Microsoft Office 里"另存为" .docx / .pptx 后重新上传。';

export async function convertToOoxml(
  buffer: Buffer,
  originalName: string
): Promise<ConvertedDoc> {
  const bin = await findSoffice();
  if (!bin) {
    throw new Error(HELP_MSG);
  }

  const lower = originalName.toLowerCase();
  let target: 'docx' | 'pptx';
  if (lower.endsWith('.doc')) target = 'docx';
  else if (lower.endsWith('.ppt')) target = 'pptx';
  else throw new Error(`不支持转换的格式：${originalName}`);

  const tmp = await mkdtemp(join(tmpdir(), 'ooxml-conv-'));
  try {
    const safeName = basename(originalName).replace(/[^\w.一-龥-]/g, '_');
    const inputPath = join(tmp, safeName);
    await writeFile(inputPath, buffer);

    logger.info('Converting via libreoffice', {
      from: lower.slice(-4),
      to: target,
      size: buffer.length,
    });

    await exec(
      bin,
      ['--headless', '--convert-to', target, '--outdir', tmp, inputPath],
      { timeout: 60_000 }
    );

    const files = await readdir(tmp);
    const out = files.find(
      (f) => f.toLowerCase().endsWith('.' + target) && f !== safeName
    );
    if (!out) throw new Error('libreoffice 未生成输出文件（请确认源文件是有效的 .doc/.ppt）');
    const outBuf = await readFile(join(tmp, out));
    return { buffer: outBuf, filename: out, kind: target };
  } finally {
    await rm(tmp, { recursive: true, force: true });
  }
}

export async function libreofficeStatus(): Promise<{ available: boolean; path: string | null }> {
  const bin = await findSoffice();
  return { available: !!bin, path: bin };
}
