import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/index.js';
import { anchorToPixelBox } from '../layout/index.js';
import type { SmartArtObject, Workbook } from '../model.js';
import { saveXlsx } from '../write/index.js';
import { loadXlsx } from './index.js';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-smartart.xlsx')),
	);

const smartArt = (wb: Workbook): SmartArtObject => {
	const found = wb.sheets[0]?.drawings.find((d) => d.kind === 'smartArt');
	if (!found || found.kind !== 'smartArt') throw new Error('no SmartArt');
	return found;
};

const DIAGRAM_PARTS = [
	'xl/diagrams/data1.xml',
	'xl/diagrams/layout1.xml',
	'xl/diagrams/quickStyle1.xml',
	'xl/diagrams/colors1.xml',
	'xl/diagrams/drawing1.xml',
];

describe('SmartArt in worksheet drawings', () => {
	it('loads an Excel SmartArt graphic with its cached drawing and node text', async () => {
		const wb = await loadXlsx(fixture());
		const art = smartArt(wb);
		expect(art.name).toBe('Process Diagram');
		expect(art.nodes.map((n) => n.text)).toEqual(['Plan', 'Build', 'Ship']);
		expect(art.layoutId).toMatch(/layout\/default$/);
		expect(art.diagram?.shapes.length).toBeGreaterThanOrEqual(3);
		expect(art.diagram?.shapes.map((s) => s.text?.paragraphs?.[0]?.runs?.[0]?.text)).toEqual(
			expect.arrayContaining(['Plan', 'Build', 'Ship']),
		);
		expect(art.notice).toMatch(/cached/);
		expect(art.sourceXml).toContain('graphicFrame');
		expect(art.anchor.from).toMatchObject({ col: 0, row: 2 });
		expect(wb.warnings.some((w) => /SmartArt/.test(w))).toBe(false);
		// Layout helpers treat it like any other anchored object.
		const box = anchorToPixelBox(wb.sheets[0]!, art.anchor);
		expect(box.w).toBeGreaterThan(0);
		expect(box.h).toBeGreaterThan(0);
	});

	it('saves an untouched SmartArt drawing byte for byte with every diagram part', async () => {
		const source = await JSZip.loadAsync(fixture());
		const saved = await JSZip.loadAsync(await saveXlsx(await loadXlsx(fixture())));
		for (const part of ['xl/drawings/drawing1.xml', ...DIAGRAM_PARTS])
			expect(await saved.file(part)?.async('string'), part).toBe(
				await source.file(part)?.async('string'),
			);
	});

	it('writes a moved SmartArt frame at the new anchor and keeps its parts and rels', async () => {
		const wb = await loadXlsx(fixture());
		const index = wb.sheets[0]!.drawings.indexOf(smartArt(wb));
		const moved = structuredClone(smartArt(wb).anchor);
		moved.from.row += 3;
		if (moved.to) moved.to.row += 3;
		createEditSession(wb).setDrawingAnchor(0, index, moved);
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		const drawing = (await zip.file('xl/drawings/drawing1.xml')?.async('string')) ?? '';
		expect(drawing).toContain('<xdr:row>5</xdr:row>');
		expect(drawing).toContain('dgm:relIds');
		const rels = (await zip.file('xl/drawings/_rels/drawing1.xml.rels')?.async('string')) ?? '';
		for (const type of ['diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors'])
			expect(rels).toContain(`relationships/${type}"`);
		expect(rels).toContain('relationships/diagramDrawing"');
		for (const part of DIAGRAM_PARTS) expect(zip.file(part), part).not.toBeNull();
		const back = smartArt(await loadXlsx(bytes));
		expect(back.anchor.from.row).toBe(5);
		expect(back.nodes.map((n) => n.text)).toEqual(['Plan', 'Build', 'Ship']);
		expect(back.diagram).toBeDefined();
	});

	it('reports a SmartArt graphic whose parts are missing instead of failing', async () => {
		const zip = await JSZip.loadAsync(fixture());
		zip.remove('xl/diagrams/drawing1.xml');
		zip.remove('xl/diagrams/data1.xml');
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const art = smartArt(wb);
		expect(art.diagram).toBeUndefined();
		expect(art.issues.some((i) => i.code === 'DIAGRAM_PART_MISSING')).toBe(true);
		expect(wb.warnings.some((w) => /SmartArt/.test(w))).toBe(true);
	});
});
