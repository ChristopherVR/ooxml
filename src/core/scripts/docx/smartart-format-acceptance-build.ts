// Creates native Word and Excel acceptance inputs from the committed producer fixtures.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { loadDocx, editDocxDiagram, formatDocxDiagram, diagramsIn } from '../../docx/index';
import { loadXlsx } from '../../xlsx/read/index';
import { saveXlsx } from '../../xlsx/write/index';
import { createEditSession } from '../../xlsx/edit/index';
import type { SmartArtObject } from '../../xlsx/model';
import {
	DIAGRAM_NATIVE_COLOR_ENTRIES,
	DIAGRAM_NATIVE_STYLE_ENTRIES,
	type DiagramFormatChange,
	type DiagramNodeEdit,
} from '../../diagram/index';
import { builtinSmartArtFormatChange } from '../../diagram/builtin/index';

const output = process.argv[2];
if (!output) throw new Error('Expected an output directory.');
mkdirSync(output, { recursive: true });
const manifest: Record<string, { texts: string[]; colors: string; style: string }> = {};
// The three model combinations, then every built-in colour definition, each with a built-in
// quick style in turn, so all 38 colour and 14 quick-style definitions (the nine 3-D ones too)
// are reopened natively.
// The `edit-*` cases first edit a node's text and add a node, so the cache Office laid out is
// regenerated around Office's own presentation points (with a bevel and a Scene quick style).
const EDITS: DiagramNodeEdit[] = [
	{ kind: 'text', nodeId: '', text: 'Beta edited' },
	{ kind: 'add', position: 'after', text: 'Zeta' },
];
const edited = <N extends { id: string }>(
	nodes: readonly N[],
	apply: (edit: DiagramNodeEdit) => void,
): void => {
	apply({ ...EDITS[0]!, nodeId: nodes[1]!.id } as DiagramNodeEdit);
	apply({ ...EDITS[1]!, nodeId: nodes.at(-1)!.id } as DiagramNodeEdit);
};
const cases: {
	name: string;
	colors: DiagramFormatChange;
	style: DiagramFormatChange;
	edit?: true;
}[] = [
	...(['3d1', '3d9'] as const).map((id) => ({
		name: `edit-${id}`,
		colors: builtinSmartArtFormatChange('colors', 'accent1_2'),
		style: builtinSmartArtFormatChange('style', id),
		edit: true as const,
	})),
	...(['flat', 'moderate', 'intense'] as const).map((style, index) => ({
		name: style,
		colors: {
			kind: 'colors' as const,
			value: (['monochromatic2', 'colorful1', 'colorful3'] as const)[index]!,
		},
		style: { kind: 'style' as const, value: style },
	})),
	...DIAGRAM_NATIVE_COLOR_ENTRIES.map((entry, index) => {
		const style = DIAGRAM_NATIVE_STYLE_ENTRIES[index % DIAGRAM_NATIVE_STYLE_ENTRIES.length]!;
		return {
			name: `native-${entry.id}-${style.id}`,
			colors: builtinSmartArtFormatChange('colors', entry.id),
			style: builtinSmartArtFormatChange('style', style.id),
		};
	}),
];
for (const { name: style, colors, style: styleChange, edit } of cases) {
	const loaded = await loadDocx(
		new Uint8Array(
			readFileSync(new URL('../../docx/__fixtures__/smartart-word.docx', import.meta.url)),
		),
	);
	const model = structuredClone(loaded.model);
	const run = model.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.diagram)!;
	let diagram = run.image!.diagram!;
	if (edit) edited(diagram.nodes, (change) => (diagram = editDocxDiagram(diagram, change).diagram));
	diagram = formatDocxDiagram(diagram, colors);
	diagram = formatDocxDiagram(diagram, styleChange);
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
	if (edit)
		edited((workbook.sheets[0]!.drawings[indexInSheet] as SmartArtObject).nodes, (change) =>
			session.editSmartArt(0, indexInSheet, change),
		);
	session.formatSmartArt(0, indexInSheet, colors);
	session.formatSmartArt(0, indexInSheet, styleChange);
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
