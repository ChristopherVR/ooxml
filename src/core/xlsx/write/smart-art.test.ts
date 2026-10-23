import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	BUILTIN_SMARTART_COLORS_ID,
	BUILTIN_SMARTART_COLORS_XML,
	BUILTIN_SMARTART_QUICK_STYLE_ID,
	BUILTIN_SMARTART_QUICK_STYLE_XML,
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
} from '../../diagram/builtin/index';
import {
	diagramNodeOutline,
	type DiagramLayoutPreset,
	type DiagramLayoutType,
} from '../../diagram/index';
import { createEditSession } from '../edit/index';
import type { SmartArtObject, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-smartart.xlsx')),
	);

const smartArtIndex = (wb: Workbook, sheet = 0): number =>
	wb.sheets[sheet]!.drawings.findIndex((d) => d.kind === 'smartArt');
const smartArt = (wb: Workbook, sheet = 0): SmartArtObject =>
	wb.sheets[sheet]!.drawings[smartArtIndex(wb, sheet)] as SmartArtObject;
const text = async (zip: JSZip, part: string) => (await zip.file(part)?.async('string')) ?? '';

const outline = (object: SmartArtObject) =>
	diagramNodeOutline(object.nodes).map((line) => `${'  '.repeat(line.level)}${line.text}`);

function layout(preset: DiagramLayoutPreset) {
	const layoutId = builtinLayoutIdForPreset(preset);
	if (!layoutId) throw new Error(`no built-in layout for ${preset}`);
	return { layoutId, layoutXml: loadBuiltinSmartArtLayoutXml(layoutId) };
}

describe('editing SmartArt in a workbook', () => {
	it('applies node commands as undoable steps and drops the stale cached drawing', async () => {
		const wb = await loadXlsx(fixture());
		const session = createEditSession(wb);
		const index = smartArtIndex(wb);
		const build = smartArt(wb).nodes[1]!.id;
		const added = session.editSmartArt(0, index, {
			kind: 'add',
			position: 'after',
			nodeId: build,
			text: 'Test',
		});
		expect(added.changed).toBe(true);
		expect(outline(smartArt(wb))).toEqual(['Plan', 'Build', 'Test', 'Ship']);
		expect(smartArt(wb).diagram).toBeUndefined();
		expect(smartArt(wb).edited).toBe(true);
		session.editSmartArt(0, index, { kind: 'demote', nodeId: added.nodeId! });
		expect(outline(smartArt(wb))).toEqual(['Plan', 'Build', '  Test', 'Ship']);
		expect(
			session.editSmartArt(0, index, { kind: 'demote', nodeId: smartArt(wb).nodes[0]!.id }),
		).toMatchObject({ changed: false });
		session.undo();
		session.undo();
		expect(outline(smartArt(wb))).toEqual(['Plan', 'Build', 'Ship']);
		expect(smartArt(wb).diagram?.shapes.length).toBeGreaterThan(0);
		expect(smartArt(wb).edited).toBeUndefined();
	});

	it('saves an edit by rewriting the data part and removing the cached drawing', async () => {
		const wb = await loadXlsx(fixture());
		const session = createEditSession(wb);
		const index = smartArtIndex(wb);
		const [plan, build, ship] = smartArt(wb).nodes.map((node) => node.id);
		session.editSmartArt(0, index, { kind: 'text', nodeId: plan!, text: 'Plan it' });
		session.editSmartArt(0, index, { kind: 'remove', nodeId: ship! });
		const added = session.editSmartArt(0, index, {
			kind: 'add',
			position: 'below',
			nodeId: build!,
			text: 'Child',
		});
		expect(added.changed).toBe(true);
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		expect(zip.file('xl/diagrams/drawing1.xml')).toBeNull();
		const rels = await text(zip, 'xl/drawings/_rels/drawing1.xml.rels');
		expect(rels).not.toContain('diagramDrawing');
		expect(rels).toContain('relationships/diagramData"');
		expect(await text(zip, '[Content_Types].xml')).not.toContain('/xl/diagrams/drawing1.xml');
		const data = await text(zip, 'xl/diagrams/data1.xml');
		expect(data).not.toContain('dataModelExt');
		expect(data).toContain('>Plan it<');
		expect(data).not.toContain('>Ship<');
		// The definitions and the frame are kept.
		const source = await JSZip.loadAsync(fixture());
		for (const part of ['layout1', 'quickStyle1', 'colors1'])
			expect(await text(zip, `xl/diagrams/${part}.xml`)).toBe(
				await text(source, `xl/diagrams/${part}.xml`),
			);
		const back = smartArt(await loadXlsx(bytes));
		expect(outline(back)).toEqual(['Plan it', 'Build', '  Child']);
		expect(back.diagram).toBeUndefined();
		expect(back.edited).toBeUndefined();
	});

	it('changes the layout: writes the new definition and names it on the data part', async () => {
		const wb = await loadXlsx(fixture());
		const session = createEditSession(wb);
		const next = layout('basicChevronProcess');
		expect(session.changeSmartArtLayout(0, smartArtIndex(wb), { ...next, family: 'process' })).toBe(
			true,
		);
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await text(zip, 'xl/diagrams/layout1.xml')).toContain(`uniqueId="${next.layoutId}"`);
		const data = await text(zip, 'xl/diagrams/data1.xml');
		expect(data).toContain(`loTypeId="${next.layoutId}"`);
		expect(data).not.toContain('type="pres"');
		const back = smartArt(await loadXlsx(await zip.generateAsync({ type: 'uint8array' })));
		expect(back.layoutId).toBe(next.layoutId);
		expect(outline(back)).toEqual(['Plan', 'Build', 'Ship']);
	});
});

describe('inserting SmartArt in a workbook', () => {
	const insertion = (
		items: string[],
		preset: DiagramLayoutPreset = 'basicBlockList',
		family: DiagramLayoutType = 'list',
	) => ({
		anchor: {
			from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
			ext: { cx: 4572000, cy: 2667000 },
		},
		...layout(preset),
		family,
		colorsId: BUILTIN_SMARTART_COLORS_ID,
		colorsXml: BUILTIN_SMARTART_COLORS_XML,
		quickStyleId: BUILTIN_SMARTART_QUICK_STYLE_ID,
		quickStyleXml: BUILTIN_SMARTART_QUICK_STYLE_XML,
		items,
	});

	it('writes the frame, four parts, relationships and content types, and reads them back', async () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);
		const index = session.insertSmartArt(0, insertion(['One', 'Two', 'Three']));
		expect(smartArt(wb).inserted).toBe(true);
		session.editSmartArt(0, index, {
			kind: 'demote',
			nodeId: smartArt(wb).nodes[2]!.id,
		});
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		const drawing = await text(zip, 'xl/drawings/drawing1.xml');
		expect(drawing).toContain('<dgm:relIds');
		expect(drawing).toMatch(/r:dm="rId\d+" r:lo="rId\d+" r:qs="rId\d+" r:cs="rId\d+"/);
		const rels = await text(zip, 'xl/drawings/_rels/drawing1.xml.rels');
		for (const type of ['diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors'])
			expect(rels).toContain(`relationships/${type}"`);
		const types = await text(zip, '[Content_Types].xml');
		expect(types).toContain(
			'PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"',
		);
		expect(types).toContain('diagramLayout+xml');
		for (const part of ['data1', 'layout1', 'quickStyle1', 'colors1'])
			expect(await text(zip, `xl/diagrams/${part}.xml`)).toMatch(/^<\?xml /);
		const back = await loadXlsx(bytes);
		const art = smartArt(back);
		expect(art.issues.filter((issue) => !issue.code.startsWith('DIAGRAM_DRAWING'))).toEqual([]);
		expect(outline(art)).toEqual(['One', 'Two', '  Three']);
		expect(art.layoutId).toBe(layout('basicBlockList').layoutId);
		expect(art.colorsId).toBe(BUILTIN_SMARTART_COLORS_ID);
		// Saving the reloaded workbook untouched keeps the parts as written.
		const again = await JSZip.loadAsync(await saveXlsx(back));
		expect(await text(again, 'xl/diagrams/data1.xml')).toBe(
			await text(zip, 'xl/diagrams/data1.xml'),
		);
	});

	it('puts a hierarchy under one top node and avoids existing part names', async () => {
		const wb = await loadXlsx(fixture());
		const session = createEditSession(wb);
		session.insertSmartArt(0, insertion(['Top', 'Left', 'Right'], 'hierarchy', 'hierarchy'));
		const inserted = wb.sheets[0]!.drawings.at(-1) as SmartArtObject;
		expect(outline(inserted)).toEqual(['Top', '  Left', '  Right']);
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		expect(zip.file('xl/diagrams/data2.xml')).not.toBeNull();
		// The existing, unedited diagram keeps its parts and cached drawing.
		const source = await JSZip.loadAsync(fixture());
		for (const part of ['data1', 'drawing1'])
			expect(await text(zip, `xl/diagrams/${part}.xml`)).toBe(
				await text(source, `xl/diagrams/${part}.xml`),
			);
		const back = await loadXlsx(bytes);
		const arts = back.sheets[0]!.drawings.filter((d) => d.kind === 'smartArt') as SmartArtObject[];
		expect(arts.map((art) => art.nodes.map((node) => node.text))).toEqual([
			['Plan', 'Build', 'Ship'],
			['Top', 'Left', 'Right'],
		]);
		session.undo();
		expect(wb.sheets[0]!.drawings.filter((d) => d.kind === 'smartArt')).toHaveLength(1);
	});
});
