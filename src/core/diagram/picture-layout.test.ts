import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from '../docx/index';
import { editDocxDiagram } from '../docx/diagram-edit';
import { previewDocxDiagram } from '../docx/ui/smartart-approximate';
import { withNativePictureLayout } from './test-support/picture-layout';
import { loadXlsx } from '../xlsx/read/index';
import { saveXlsx } from '../xlsx/write/index';
import { createEditSession } from '../xlsx/edit/index';
import type { SmartArtObject } from '../xlsx/model';
import { loadBuiltinSmartArtLayoutXml } from './builtin/index';

it('retains picture identity, crops and effects through Excel edits and save', async () => {
	const bytes = await withNativePictureLayout(
		readFileSync(new URL('../xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url)),
		'xl',
	);
	const workbook = await loadXlsx(bytes);
	const index = workbook.sheets[0]!.drawings.findIndex((item) => item.kind === 'smartArt');
	const object = workbook.sheets[0]!.drawings[index] as SmartArtObject;
	const session = createEditSession(workbook);
	session.editSmartArt(0, index, {
		kind: 'text',
		nodeId: object.nodes[0]!.id,
		text: 'Excel image caption',
	});
	const reopened = await loadXlsx(await saveXlsx(workbook));
	const result = reopened.sheets[0]!.drawings[index] as SmartArtObject;
	expect(result.diagram?.shapes.filter((shape) => shape.fill?.kind === 'picture')).toHaveLength(3);
	expect(result.layoutSource?.pictures).toHaveLength(3);
	expect(
		result.layoutSource?.pictures?.every(
			(source) =>
				source.fill.crop?.cropRight === 0.25 && source.fill.sourceXml?.includes('alphaModFix'),
		),
	).toBe(true);
});

it('retains dormant image associations through a non-picture layout, reopen and picture-grid switch', async () => {
	const bytes = await withNativePictureLayout(
		readFileSync(new URL('../xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url)),
		'xl',
	);
	let workbook = await loadXlsx(bytes);
	const index = workbook.sheets[0]!.drawings.findIndex((item) => item.kind === 'smartArt');
	let session = createEditSession(workbook);
	const plain = 'urn:microsoft.com/office/officeart/2005/8/layout/default';
	session.changeSmartArtLayout(0, index, {
		layoutId: plain,
		layoutXml: loadBuiltinSmartArtLayoutXml(plain),
		family: 'list',
	});
	workbook = await loadXlsx(await saveXlsx(workbook));
	let object = workbook.sheets[0]!.drawings[index] as SmartArtObject;
	expect(object.layoutSource?.pictures).toHaveLength(3);
	expect(object.diagram).toBeUndefined();
	session = createEditSession(workbook);
	const grid = 'urn:microsoft.com/office/officeart/2008/layout/PictureGrid';
	session.changeSmartArtLayout(0, index, {
		layoutId: grid,
		layoutXml: loadBuiltinSmartArtLayoutXml(grid),
		family: 'matrix',
	});
	session.editSmartArt(0, index, { kind: 'move', nodeId: object.nodes[0]!.id, direction: 1 });
	session.editSmartArt(0, index, { kind: 'remove', nodeId: object.nodes[2]!.id });
	workbook = await loadXlsx(await saveXlsx(workbook));
	object = workbook.sheets[0]!.drawings[index] as SmartArtObject;
	expect(object.layoutSource?.pictures).toHaveLength(2);
	expect(object.diagram?.shapes.filter((shape) => shape.fill?.kind === 'picture')).toHaveLength(2);
});

it('reflows a native Picture Accent List without assigning pictures to sibling text roles', async () => {
	const bytes = await withNativePictureLayout(
		readFileSync(new URL('../docx/__fixtures__/smartart.docx', import.meta.url)),
		'word',
	);
	const loaded = await loadDocx(bytes);
	const image = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? block.runs.flatMap((run) => (run.image?.diagram ? [run.image] : []))
			: [],
	)[0]!;
	const original = image.diagram!;
	expect(original.layoutSource?.pictures?.length).toBeGreaterThan(0);
	const edited = editDocxDiagram(original, {
		kind: 'text',
		nodeId: original.nodes[0]!.id,
		text: 'Edited picture caption',
	}).diagram;
	const preview = previewDocxDiagram(edited, loaded.model.theme);
	expect(preview.mode).toBe('interpreted');
	expect(preview.issues).toEqual([]);
	expect(preview.drawing?.extentEmu).toEqual(edited.extentEmu);
	const panels = preview.drawing?.shapes.filter((shape) => shape.sourceStyleLabel === 'node1');
	expect(panels?.map((shape) => shape.fill)).toEqual(
		Array(3).fill({ kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } }),
	);
	expect(preview.drawing?.shapes.filter((shape) => shape.fill?.kind === 'picture')).toHaveLength(3);
	expect(
		preview.drawing?.shapes
			.filter((shape) =>
				shape.text?.paragraphs.some((paragraph) =>
					paragraph.runs.some((run) => run.text.includes('Edited')),
				),
			)
			.every((shape) => shape.fill?.kind !== 'picture'),
	).toBe(true);
	image.diagram = edited;
	const saved = await loaded.save(loaded.model);
	const again = await loadDocx(saved);
	const reread = again.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? block.runs.flatMap((run) => (run.image?.diagram ? [run.image.diagram] : []))
			: [],
	)[0]!;
	expect(reread.drawing?.shapes.filter((shape) => shape.fill?.kind === 'picture')).toHaveLength(3);
	expect(
		reread.layoutSource?.pictures?.every((source) => source.fill.crop?.cropLeft === 0.125),
	).toBe(true);
});
