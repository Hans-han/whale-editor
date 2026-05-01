const LATIN1_MOJIBAKE_RE = /[\u00c0-\u00ff]/;
const CJK_RE = /[\u3400-\u9fff]/;
const REPLACEMENT_RE = /\uFFFD/;

const CP1252_BYTES = new Map<string, number>([
  ['€', 0x80],
  ['‚', 0x82],
  ['ƒ', 0x83],
  ['„', 0x84],
  ['…', 0x85],
  ['†', 0x86],
  ['‡', 0x87],
  ['ˆ', 0x88],
  ['‰', 0x89],
  ['Š', 0x8a],
  ['‹', 0x8b],
  ['Œ', 0x8c],
  ['Ž', 0x8e],
  ['‘', 0x91],
  ['’', 0x92],
  ['“', 0x93],
  ['”', 0x94],
  ['•', 0x95],
  ['–', 0x96],
  ['—', 0x97],
  ['˜', 0x98],
  ['™', 0x99],
  ['š', 0x9a],
  ['›', 0x9b],
  ['œ', 0x9c],
  ['ž', 0x9e],
  ['Ÿ', 0x9f],
]);

function mojibakeBytes(text: string): Buffer | null {
  const bytes: number[] = [];
  for (const char of text) {
    const code = char.codePointAt(0);
    if (code === undefined) return null;
    if (code <= 0xff) {
      bytes.push(code);
      continue;
    }
    const cp1252 = CP1252_BYTES.get(char);
    if (cp1252 === undefined) return null;
    bytes.push(cp1252);
  }
  return Buffer.from(bytes);
}

function decodeMojibakeOnce(text: string): string | null {
  const bytes = mojibakeBytes(text);
  if (!bytes) return null;
  return bytes.toString('utf8');
}

function filenameScore(text: string): number {
  const cjk = [...text].filter((char) => CJK_RE.test(char)).length;
  const mojibake = [...text].filter((char) => LATIN1_MOJIBAKE_RE.test(char) || CP1252_BYTES.has(char)).length;
  const replacement = [...text].filter((char) => REPLACEMENT_RE.test(char)).length;
  return cjk * 8 - mojibake * 2 - replacement * 50;
}

export function normalizeUploadFilename(name: string): string {
  if (!name || !LATIN1_MOJIBAKE_RE.test(name)) return name;

  let best = name;
  let bestScore = filenameScore(name);
  let current = name;

  for (let i = 0; i < 3; i++) {
    const decoded = decodeMojibakeOnce(current);
    if (!decoded || decoded === current) break;

    const score = filenameScore(decoded);
    if (score > bestScore && !REPLACEMENT_RE.test(decoded)) {
      best = decoded;
      bestScore = score;
    }
    current = decoded;
  }

  return CJK_RE.test(best) ? best : name;
}
