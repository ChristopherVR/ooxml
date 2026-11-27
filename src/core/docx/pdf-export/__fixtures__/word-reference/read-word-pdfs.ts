// Reads the PDFs `src/core/scripts/docx-word-pdf-reference.ps1` saved with Word into
// `word-lines.json`: per reference document, Word's page count and the lines of every page.
// `bun src/core/docx/pdf-export/__fixtures__/word-reference/read-word-pdfs.ts <pdf folder>`
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pdfLines } from './pdf-lines';

const folder = process.argv[2];
if (!folder) throw new Error('Pass the folder the Word PDFs were saved in.');
const result: Record<string, string[][]> = {};
for (const name of [
	'footnotes',
	'header-footer',
	'landscape',
	'letter',
	'table-header',
	'table-header-no-widow-control',
])
	result[name] = await pdfLines(new Uint8Array(readFileSync(join(folder, `${name}.pdf`))));
writeFileSync(
	new URL('./word-lines.json', import.meta.url),
	`${JSON.stringify(result, null, '\t')}\n`,
);
