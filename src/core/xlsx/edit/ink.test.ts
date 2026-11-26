import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inkStrokeFromSamples } from '../../drawingml/inkml/edit';
import type { InkObject, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { addInk, setInk } from './ink';
import { testContext } from './test-context';

// Excel 16's own saves: ink pasted from PowerPoint (with its fallback picture), and ink Excel
// re-saved from markup without one (`xlsx/__fixtures__/README.md`).
const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name)));

const inkOf = (wb: Workbook): InkObject => {
	const ink = wb.sheets[0]?.drawings.find((drawing) => drawing.kind === 'ink');
	if (ink?.kind !== 'ink') throw new Error('no ink');
	return ink;
};

describe('Excel ink', () => {
	it.each(['excel-ink.xlsx', 'excel-ink-fallback.xlsx'])(
		'reads the strokes of %s',
		async (name) => {
			const ink = inkOf(await loadXlsx(fixture(name)));
			expect(ink.partName).toBe('xl/ink/ink1.xml');
			expect(ink.strokes).toHaveLength(4);
			expect(ink.strokes[0]?.color).toBe('#E81123');
			expect(ink.width).toBeGreaterThan(100);
		},
	);

	it('keeps untouched ink, its fallback picture and its part byte for byte', async () => {
		const original = fixture('excel-ink-fallback.xlsx');
		const wb = await loadXlsx(original);
		const ctx = testContext(wb);
		// An edit elsewhere on the sheet regenerates the drawing part.
		addInk(
			ctx,
			0,
			{ from: { row: 12, col: 1, rowOffset: 0, colOffset: 0 }, ext: { cx: 952500, cy: 476250 } },
			{
				strokes: [
					inkStrokeFromSamples(
						[
							{ x: 1, y: 1 },
							{ x: 90, y: 40 },
						],
						{ kind: 'pen', color: '#000000', width: 2 },
					)!,
				],
				width: 100,
				height: 50,
			},
		);
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		const source = await JSZip.loadAsync(original);
		expect(await saved.file('xl/ink/ink1.xml')!.async('string')).toBe(
			await source.file('xl/ink/ink1.xml')!.async('string'),
		);
		const drawing = await saved.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing).toContain('<mc:Fallback>');
		expect(drawing.match(/<xdr:contentPart/g)).toHaveLength(2);
		expect(saved.file('xl/ink/ink2.xml')).toBeTruthy();
		const types = await saved.file('[Content_Types].xml')!.async('string');
		expect(types).toContain('PartName="/xl/ink/ink2.xml" ContentType="application/inkml+xml"');
	});

	it('saves new ink and erased strokes, and reads them back', async () => {
		const wb = await loadXlsx(fixture('excel-ink.xlsx'));
		const ctx = testContext(wb);
		const ink = inkOf(wb);
		const kept = ink.strokes.slice(2);
		setInk(ctx, 0, 0, ink.anchor, { strokes: kept, width: ink.width, height: ink.height });
		const reread = inkOf(await loadXlsx(await saveXlsx(wb)));
		expect(reread.strokes.map((stroke) => stroke.path)).toEqual(kept.map((stroke) => stroke.path));
		ctx.undo();
		expect(inkOf(wb).strokes).toHaveLength(4);
	});

	it('deletes ink the eraser left empty, with its part and relationship', async () => {
		const wb = await loadXlsx(fixture('excel-ink.xlsx'));
		const ctx = testContext(wb);
		const ink = inkOf(wb);
		addInk(ctx, 0, ink.anchor, {
			strokes: ink.strokes.slice(0, 1),
			width: ink.width,
			height: ink.height,
		});
		setInk(ctx, 0, 0, ink.anchor, { strokes: [], width: ink.width, height: ink.height });
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const rels = await zip.file('xl/drawings/_rels/drawing1.xml.rels')!.async('string');
		expect(rels.match(/customXml/g)).toHaveLength(1);
		const reread = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(reread.sheets[0]?.drawings.filter((drawing) => drawing.kind === 'ink')).toHaveLength(1);
	});
});
