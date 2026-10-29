// Builds the documents `docx-word-corpus-reopen.ps1` reopens in Word. Run with bun from src/core:
//   bun scripts/docx-word-corpus-reopen-build.ts <output folder>
// Each Word-authored corpus document (`docx/__fixtures__/word-corpus/`) is loaded with the core,
// given one trivial edit (a marker word appended to the last plain run of the first body paragraph
// that has one) and saved with `saveDocx`. `manifest.json` lists the marker and the page count
// Word recorded for the original, which the script compares after Word reopens the saved file.
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { loadDocx } from '../docx/index';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/docx-word-corpus-reopen-build.ts <output folder>');
mkdirSync(out, { recursive: true });
const corpus = join(import.meta.dir, '..', 'docx', '__fixtures__', 'word-corpus');
const marker = ' EDITED';
const manifest: Record<string, { marker: string; pageCount: number; edited: boolean }> = {};

for (const file of readdirSync(corpus).filter((f) => f.endsWith('.docx'))) {
	const name = file.replace(/\.docx$/, '');
	const facts = JSON.parse(readFileSync(join(corpus, `${name}.facts.json`), 'utf8')) as {
		pageCount: number;
	};
	const loaded = await loadDocx(new Uint8Array(readFileSync(join(corpus, file))));
	const model = structuredClone(loaded.model);
	let edited = false;
	for (const block of model.blocks) {
		if (block.type !== 'paragraph') continue;
		const run = [...block.runs]
			.reverse()
			.find(
				(r) => r.text.trim() && !r.equation && r.text === r.text.replace(/[\u0000-\u001f]/g, ''),
			);
		if (!run) continue;
		run.text += marker;
		edited = true;
		break;
	}
	manifest[file] = { marker: marker.trim(), pageCount: facts.pageCount, edited };
	writeFileSync(join(out, file), await loaded.save(model));
}
writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} documents to ${out}`);
