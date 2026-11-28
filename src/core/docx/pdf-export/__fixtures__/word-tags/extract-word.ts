// Reads Word 16's Save as PDF of the issue #415 and #427 documents into `word-tags.json`: every
// line's baseline, size, text and glyph positions. `bun extract-word.ts` reads the PDFs (issue-415-word.pdf, issue-427-bug-word.pdf, issue-427-control-word.pdf) from next to it; they are not committed.
import { readFileSync, writeFileSync } from 'node:fs';
import { pdfGlyphLines } from '../word-sizes/pdf-glyphs';

const round = (value: number) => Math.round(value * 1000) / 1000;
const files = {
	'issue-415': 'issue-415-word.pdf',
	'issue-427-bug': 'issue-427-bug-word.pdf',
	'issue-427-control': 'issue-427-control-word.pdf',
};
const result: Record<string, unknown[]> = {};
for (const [name, file] of Object.entries(files)) {
	const lines = await pdfGlyphLines(new Uint8Array(readFileSync(new URL(file, import.meta.url))));
	result[name] = lines.map((line) => ({
		page: line.page,
		y: round(line.y),
		size: round(line.size),
		font: line.font,
		text: line.text,
		x: line.x.map(round),
	}));
}
const json = JSON.stringify(result)
	.replace(/\{"page"/g, '\n\t{"page"')
	.replace(/\]\}$/, ']\n}\n');
writeFileSync(new URL('./word-tags.json', import.meta.url), json);
