// Builds the documents `docx-word-acceptance.ps1` opens in Word. Run with bun from src/core:
//   bun scripts/docx-acceptance-build.ts <output folder>
// Every file is written by `saveDocx` through `loadDocx(...).save`: the committed SmartArt fixtures
// saved untouched, saved after node edits (text, add, remove, move, promote, demote), and with
// SmartArt inserted from Insert SmartArt presets. `manifest.json` lists, per file, the node texts
// Word must report for each SmartArt (`SmartArt.AllNodes`, tree order).
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
	BUILTIN_SMARTART_COLORS_ID,
	BUILTIN_SMARTART_COLORS_XML,
	BUILTIN_SMARTART_QUICK_STYLE_ID,
	BUILTIN_SMARTART_QUICK_STYLE_XML,
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
} from '../diagram/builtin/index';
import { DIAGRAM_INSERT_PRESETS } from '../diagram/catalogue';
import type { DiagramLayoutPreset } from '../diagram/model/data';
import {
	createDocxDiagram,
	diagramsIn,
	editDocxDiagram,
	insertedDiagramPartNames,
	loadDocx,
	type DocumentModel,
	type DocxDiagram,
	type PendingMediaPart,
} from '../docx/index';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/docx-acceptance-build.ts <output folder>');
mkdirSync(out, { recursive: true });
const fixtures = join(import.meta.dir, '..', 'docx', '__fixtures__');
const manifest: Record<string, string[][]> = {};
const xmlPart = (xml: string): PendingMediaPart => ({
	bytes: new TextEncoder().encode(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${xml}`,
	),
	contentType: 'application/xml',
});

const diagramRuns = (model: DocumentModel) => {
	const runs = [];
	for (const block of model.blocks)
		if (block.type === 'paragraph')
			for (const run of block.runs) if (run.image?.diagram) runs.push(run.image);
	return runs;
};

/** Word's AllNodes order is the tree order the model keeps; empty nodes report ''. */
const expected = (model: DocumentModel): string[][] =>
	diagramsIn(model.blocks).map((diagram) => diagram.nodes.map((node) => node.text));

async function write(
	name: string,
	source: string,
	edit?: (model: DocumentModel, media: Map<string, PendingMediaPart>) => void,
) {
	const loaded = await loadDocx(new Uint8Array(readFileSync(join(fixtures, source))));
	const model = structuredClone(loaded.model);
	const media = new Map<string, PendingMediaPart>();
	edit?.(model, media);
	writeFileSync(join(out, name), await loaded.save(model, media));
	manifest[name] = expected(model);
}

function editFirst(model: DocumentModel, change: (diagram: DocxDiagram) => DocxDiagram): void {
	const image = diagramRuns(model)[0];
	if (!image?.diagram) throw new Error('no diagram');
	image.diagram = change(image.diagram);
}

const edits = (diagram: DocxDiagram): DocxDiagram => {
	const [a, b, c, d] = diagram.nodes;
	let next = editDocxDiagram(diagram, { kind: 'text', nodeId: b!.id, text: 'Beta edited' }).diagram;
	next = editDocxDiagram(next, {
		kind: 'add',
		position: 'after',
		nodeId: a!.id,
		text: 'After Alpha',
	}).diagram;
	next = editDocxDiagram(next, {
		kind: 'add',
		position: 'below',
		nodeId: b!.id,
		text: 'Under Beta',
	}).diagram;
	next = editDocxDiagram(next, { kind: 'remove', nodeId: c!.id }).diagram;
	next = editDocxDiagram(next, { kind: 'move', nodeId: d!.id, direction: -1 }).diagram;
	next = editDocxDiagram(next, { kind: 'add', position: 'before', nodeId: a!.id }).diagram;
	return next;
};

let suffix = 0;
function insert(
	model: DocumentModel,
	media: Map<string, PendingMediaPart>,
	preset: DiagramLayoutPreset,
) {
	const entry = DIAGRAM_INSERT_PRESETS.find((candidate) => candidate.layout === preset)!;
	const layoutId = builtinLayoutIdForPreset(preset)!;
	const partSuffix = `-acc${++suffix}`;
	const diagram = createDocxDiagram({
		layoutId,
		family: entry.family,
		colorsId: BUILTIN_SMARTART_COLORS_ID,
		quickStyleId: BUILTIN_SMARTART_QUICK_STYLE_ID,
		items: entry.defaultItems,
		widthPx: 576,
		heightPx: 336,
		partSuffix,
	});
	const names = insertedDiagramPartNames(partSuffix);
	media.set(names.layout, xmlPart(loadBuiltinSmartArtLayoutXml(layoutId)));
	media.set(names.colors, xmlPart(BUILTIN_SMARTART_COLORS_XML));
	media.set(names.quickStyle, xmlPart(BUILTIN_SMARTART_QUICK_STYLE_XML));
	model.blocks.push({
		type: 'paragraph',
		id: `acc-${suffix}`,
		runs: [
			{
				text: '',
				image: {
					relId: '',
					partName: '',
					contentType: 'application/octet-stream',
					widthPx: 576,
					heightPx: 336,
					unsupported: 'SmartArt',
					diagram,
				},
			},
		],
	});
}

await write('01-word-untouched.docx', 'smartart-word.docx');
await write('02-word-edited.docx', 'smartart-word.docx', (model) => editFirst(model, edits));
await write('03-word-promote-demote.docx', 'smartart-word.docx', (model) =>
	editFirst(model, (diagram) => {
		const [, b, c] = diagram.nodes;
		let next = editDocxDiagram(diagram, { kind: 'demote', nodeId: b!.id }).diagram;
		next = editDocxDiagram(next, { kind: 'demote', nodeId: c!.id }).diagram;
		next = editDocxDiagram(next, { kind: 'promote', nodeId: c!.id }).diagram;
		return next;
	}),
);
await write('04-pptx-derived-edited.docx', 'smartart.docx', (model) =>
	editFirst(model, (diagram) => {
		const first = diagram.nodes.find((node) => node.text)!;
		const renamed = editDocxDiagram(diagram, { kind: 'text', nodeId: first.id, text: 'Renamed' });
		return editDocxDiagram(renamed.diagram, {
			kind: 'add',
			position: 'below',
			nodeId: first.id,
			text: 'Added report',
		}).diagram;
	}),
);
await write('05-inserted.docx', 'smartart-word.docx', (model, media) => {
	for (const preset of [
		'basicBlockList',
		'hierarchy',
		'basicCycle',
		'basicChevronProcess',
	] as const)
		insert(model, media, preset);
});
await write('06-inserted-then-edited.docx', 'smartart-word.docx', (model, media) => {
	insert(model, media, 'verticalBlockList');
	const image = diagramRuns(model).at(-1)!;
	const [first, second] = image.diagram!.nodes;
	let next = editDocxDiagram(image.diagram!, { kind: 'demote', nodeId: second!.id }).diagram;
	next = editDocxDiagram(next, { kind: 'text', nodeId: first!.id, text: 'Edited block' }).diagram;
	image.diagram = next;
});
writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, '\t'));
console.log(`wrote ${Object.keys(manifest).length} documents to ${out}`);
