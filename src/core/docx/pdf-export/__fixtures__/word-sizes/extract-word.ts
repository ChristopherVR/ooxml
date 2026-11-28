// Reads Word 16's Save as PDF of `font-sizes.docx` and `glyph-positions.docx` (the files attached
// to issue #382) and of `line-pitch.docx` and `line-pitch-letter.docx` (single-spaced paragraphs
// in Calibri, Times New Roman and Arial at 9 to 23 pt, on A4 with a 1200 twip top margin and on
// Letter with 1000) into `word-glyphs.json`: every line's baseline, drawn size and glyph positions.
// `bun src/core/docx/pdf-export/__fixtures__/word-sizes/extract-word.ts <folder with the PDFs>`
// reads `<name>-word.pdf` for each document. To keep the file small, the glyph positions of
// `glyph-positions.docx` are kept for every third line only (its baselines and sizes for all).
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pdfGlyphLines } from './pdf-glyphs';

const folder = process.argv[2];
if (!folder) throw new Error('Pass the folder the Word PDFs were saved in.');
const round = (value: number) => Math.round(value * 1000) / 1000;
const result: Record<string, unknown[]> = {};
for (const name of ['font-sizes', 'glyph-positions', 'line-pitch', 'line-pitch-letter']) {
	const lines = await pdfGlyphLines(new Uint8Array(readFileSync(join(folder, `${name}-word.pdf`))));
	result[name] = lines.map((line, index) => ({
		page: line.page,
		y: round(line.y),
		size: round(line.size),
		font: line.font,
		text: line.text,
		...(name !== 'glyph-positions' || index % 3 === 0 ? { x: line.x.map(round) } : {}),
	}));
}
const json = JSON.stringify(result)
	.replace(/\{"page"/g, '\n\t{"page"')
	.replace(/\]\}$/, ']\n}\n');
writeFileSync(new URL('./word-glyphs.json', import.meta.url), json);
