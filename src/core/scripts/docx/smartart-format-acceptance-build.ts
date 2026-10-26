// Creates native Word and Excel acceptance inputs from the committed producer fixtures.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { loadDocx, formatDocxDiagram, diagramsIn } from '../../docx/index';
import { loadXlsx } from '../../xlsx/read/index';
import { saveXlsx } from '../../xlsx/write/index';
import { createEditSession } from '../../xlsx/edit/index';
import type { SmartArtObject } from '../../xlsx/model';

const output = process.argv[2];
if (!output) throw new Error('Expected an output directory.');
mkdirSync(output, { recursive: true });
const manifest: Record<string, { texts: string[]; colors: string; style: string }> = {};
for (const [index, style] of (['flat', 'moderate', 'intense'] as const).entries()) {
	const colors = (['monochromatic2', 'colorful1', 'colorful3'] as const)[index]!;
	const loaded = await loadDocx(
		new Uint8Array(
			readFileSync(new URL('../../docx/__fixtures__/smartart-word.docx', import.meta.url)),
		),
	);
	const model = structuredClone(loaded.model);
	const run = model.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.diagram)!;
	let diagram = formatDocxDiagram(run.image!.diagram!, { kind: 'colors', value: colors });
	diagram = formatDocxDiagram(diagram, { kind: 'style', value: style });
	run.image!.diagram = diagram;
	const wordName = `word-${style}.docx`;
	const wordBytes = await loaded.save(model);
	if (diagramsIn((await loadDocx(wordBytes)).model.blocks)[0]?.rendering !== 'generated-drawing')
		throw new Error('The Word fixture has no verified generated cache.');
	writeFileSync(join(output, wordName), wordBytes);
	manifest[wordName] = {
		texts: diagram.nodes.map((node) => node.text),
		colors: diagram.colorsId!,
		style: diagram.quickStyleId!,
	};
	const workbook = await loadXlsx(
		new Uint8Array(
			readFileSync(new URL('../../xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url)),
		),
	);
	const indexInSheet = workbook.sheets[0]!.drawings.findIndex((shape) => shape.kind === 'smartArt');
	const session = createEditSession(workbook);
	session.formatSmartArt(0, indexInSheet, { kind: 'colors', value: colors });
	session.formatSmartArt(0, indexInSheet, { kind: 'style', value: style });
	const art = workbook.sheets[0]!.drawings[indexInSheet] as SmartArtObject;
	const excelName = `excel-${style}.xlsx`;
	const excelBytes = await saveXlsx(workbook);
	if (
		!(await loadXlsx(excelBytes)).sheets[0]!.drawings.some(
			(shape) => shape.kind === 'smartArt' && shape.cacheGenerated,
		)
	)
		throw new Error('The Excel fixture has no verified generated cache.');
	writeFileSync(join(output, excelName), excelBytes);
	manifest[excelName] = {
		texts: art.nodes.map((node) => node.text),
		colors: art.colorsId!,
		style: art.quickStyleId!,
	};
}
writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest, null, 2));
console.log(`Wrote ${Object.keys(manifest).length} native formatting inputs to ${output}`);
