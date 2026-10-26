import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { diagramsIn, loadDocx, formatDocxDiagram } from '../../docx/index';
import { loadXlsx } from '../../xlsx/read/index';
import { saveXlsx } from '../../xlsx/write/index';
import { createEditSession } from '../../xlsx/edit/index';
import type { SmartArtObject } from '../../xlsx/model';
import { previewDiagramLayout } from '../layout-preview';
import { builtinLayoutIdForPreset, loadBuiltinSmartArtLayoutXml } from '../builtin';
import { formatDiagramSource } from './format';

const word = () =>
	new Uint8Array(
		readFileSync(new URL('../../docx/__fixtures__/smartart-word.docx', import.meta.url)),
	);
const excel = () =>
	new Uint8Array(
		readFileSync(new URL('../../xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url)),
	);

it('saves Word styles and colors with original text/parts intact and a fresh uncached preview', async () => {
	const loaded = await loadDocx(word());
	const edited = structuredClone(loaded.model);
	const run = edited.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.diagram)!;
	const original = run.image!.diagram!;
	let next = formatDocxDiagram(original, { kind: 'colors', value: 'monochromatic2' });
	next = formatDocxDiagram(next, { kind: 'style', value: 'moderate' });
	run.image!.diagram = next;
	expect(next.nodes).toEqual(original.nodes);
	expect(next.drawing).toBeUndefined();
	expect(formatDocxDiagram(next, { kind: 'style', value: 'moderate' })).toBe(next);
	const bytes = await loaded.save(edited);
	const reopened = diagramsIn((await loadDocx(bytes)).model.blocks)[0]!;
	expect(reopened.rendering).toBe('generated-drawing');
	expect(reopened.drawing?.shapes.length).toBeGreaterThan(0);
	expect(reopened).toMatchObject({
		colorsId: 'urn:ooxml-core/colors/monochromatic2',
		quickStyleId: 'urn:ooxml-core/style/moderate',
	});
	expect(reopened.nodes).toEqual(original.nodes);
	const before = await JSZip.loadAsync(word());
	const after = await JSZip.loadAsync(bytes);
	expect(await after.file(original.parts.layout!.partName)!.async('string')).toBe(
		await before.file(original.parts.layout!.partName)!.async('string'),
	);
	expect(await after.file('word/theme/theme1.xml')!.async('string')).toBe(
		await before.file('word/theme/theme1.xml')!.async('string'),
	);
	expect(reopened.layoutSource?.colorsXml).toContain('accent2');
});

it('makes Excel style/color choices undoable and saves their own definitions after reload', async () => {
	const workbook = await loadXlsx(excel());
	const session = createEditSession(workbook);
	const index = workbook.sheets[0]!.drawings.findIndex((shape) => shape.kind === 'smartArt');
	const art = () => workbook.sheets[0]!.drawings[index] as SmartArtObject;
	const original = structuredClone(art());
	expect(session.formatSmartArt(0, index, { kind: 'colors', value: 'colorful3' })).toBe(true);
	expect(session.formatSmartArt(0, index, { kind: 'style', value: 'intense' })).toBe(true);
	session.undo();
	expect(art().quickStyleId).toBe(original.quickStyleId);
	session.redo();
	expect(art().quickStyleId).toBe('urn:ooxml-core/style/intense');
	const saved = await saveXlsx(workbook);
	const back = await loadXlsx(saved);
	const object = back.sheets[0]!.drawings.find(
		(shape) => shape.kind === 'smartArt',
	) as SmartArtObject;
	expect(object).toMatchObject({
		colorsId: 'urn:ooxml-core/colors/colorful3',
		quickStyleId: 'urn:ooxml-core/style/intense',
	});
	expect(object.nodes).toEqual(original.nodes);
	expect(object.layoutSource?.quickStyleXml).toContain('urn:ooxml-core/style/intense');
	expect(object.diagram?.shapes.length).toBeGreaterThan(0);
	expect(object.cacheGenerated).toBe(true);
});

it('keeps an explicit node fill above a changed global palette in the interpreted preview', () => {
	const dataXml =
		'<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="a"><dgm:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></dgm:spPr></dgm:pt></dgm:ptLst></dgm:dataModel>';
	const source = formatDiagramSource(
		{
			dataXml,
			layoutXml: loadBuiltinSmartArtLayoutXml(builtinLayoutIdForPreset('basicBlockList')!),
		},
		{ kind: 'colors', value: 'monochromatic2' },
	).source;
	const input = {
		extentEmu: { width: 600 * 9525, height: 300 * 9525 },
		nodes: [{ id: 'a', text: 'Alpha' }],
		source,
	};
	const preview = previewDiagramLayout(input);
	expect(preview.mode).toBe('interpreted');
	expect(preview.drawing!.shapes.find((shape) => shape.text?.text === 'Alpha')!.fill).toEqual({
		kind: 'solid',
		color: { kind: 'srgb', value: 'FF0000', transforms: [] },
	});
	expect(source.dataXml).toBe(dataXml);
});
