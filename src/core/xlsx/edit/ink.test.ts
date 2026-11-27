import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inkStrokeFromSamples } from '../../drawingml/inkml/edit';
import type { InkObject, Workbook } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { addInk, setInk } from './ink';
import { createEditSession } from './session';
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

	it('drops the stale fallback picture of ink whose strokes changed', async () => {
		const wb = await loadXlsx(fixture('excel-ink-fallback.xlsx'));
		const ink = inkOf(wb);
		setInk(testContext(wb), 0, 0, ink.anchor, {
			strokes: ink.strokes.slice(1),
			width: ink.width,
			height: ink.height,
		});
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing).not.toContain('Fallback');
		expect(drawing).toContain('<xdr:contentPart');
		expect(
			inkOf(await loadXlsx(await zip.generateAsync({ type: 'uint8array' }))).strokes,
		).toHaveLength(3);
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

	const newStroke = () =>
		inkStrokeFromSamples(
			[
				{ x: 1, y: 1 },
				{ x: 90, y: 40 },
			],
			{ kind: 'pen', color: '#000000', width: 2 },
		)!;
	const otherAnchor = {
		from: { row: 12, col: 1, rowOffset: 0, colOffset: 0 },
		ext: { cx: 952500, cy: 476250 },
	};

	it('gives a copy of ink a part of its own', async () => {
		const wb = await loadXlsx(fixture('excel-ink.xlsx'));
		const copy = { ...structuredClone(inkOf(wb)), anchor: otherAnchor } as InkObject;
		wb.sheets[0]!.drawings.push(copy);
		const reread = await loadXlsx(await saveXlsx(wb));
		const inks = reread.sheets[0]!.drawings.filter(
			(drawing): drawing is InkObject => drawing.kind === 'ink',
		);
		expect(inks.map((ink) => ink.strokes.length)).toEqual([4, 4]);
		expect(new Set(inks.map((ink) => ink.partName)).size).toBe(2);
	});

	it('keeps the part of ink whose edited copy is written first', async () => {
		const wb = await loadXlsx(fixture('excel-ink.xlsx'));
		const original = inkOf(wb);
		const copy = { ...structuredClone(original), anchor: otherAnchor } as InkObject;
		copy.strokes = copy.strokes.slice(2);
		// The copy comes first, as on a duplicated sheet ordered before the original.
		wb.sheets[0]!.drawings.unshift(copy);
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		const source = await JSZip.loadAsync(fixture('excel-ink.xlsx'));
		expect(await zip.file('xl/ink/ink1.xml')!.async('string')).toBe(
			await source.file('xl/ink/ink1.xml')!.async('string'),
		);
		const inks = (await loadXlsx(bytes)).sheets[0]!.drawings.filter(
			(drawing): drawing is InkObject => drawing.kind === 'ink',
		);
		expect(inks.map((ink) => ink.strokes.length)).toEqual([2, 4]);
	});

	it('never rewrites a part read only in part', async () => {
		const zip = await JSZip.loadAsync(fixture('excel-ink.xlsx'));
		const big = Array.from({ length: 5000 }, (_, i) => `${i} ${i}`).join(',');
		const xml = (await zip.file('xl/ink/ink1.xml')!.async('string')).replace(
			/<\/(\w+:)?ink>\s*$/u,
			(end, prefix = '') => `<${prefix}trace>${big}</${prefix}trace>${end}`,
		);
		expect(xml).toContain(big.slice(0, 20));
		zip.file('xl/ink/ink1.xml', xml);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const ink = inkOf(wb);
		expect(ink.truncated).toBe(true);
		setInk(testContext(wb), 0, 0, ink.anchor, {
			strokes: ink.strokes.slice(1),
			width: ink.width,
			height: ink.height,
		});
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		expect(await saved.file('xl/ink/ink1.xml')!.async('string')).toBe(xml);
	});

	it('keeps ink whose part cannot be found as it was written', async () => {
		const zip = await JSZip.loadAsync(fixture('excel-ink.xlsx'));
		zip.file(
			'xl/drawings/_rels/drawing1.xml.rels',
			'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
		);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const [drawing] = wb.sheets[0]!.drawings;
		expect(drawing?.kind).toBe('unsupported');
		addInk(testContext(wb), 0, otherAnchor, { strokes: [newStroke()], width: 100, height: 50 });
		const saved = await JSZip.loadAsync(await saveXlsx(wb));
		const out = await saved.file('xl/drawings/drawing1.xml')!.async('string');
		expect(out.match(/<xdr:contentPart/g)).toHaveLength(2);
	});

	it('makes the edits of one eraser gesture one undo step', async () => {
		const wb = await loadXlsx(fixture('excel-ink.xlsx'));
		const session = createEditSession(wb);
		const ink = inkOf(wb);
		const size = { width: ink.width, height: ink.height };
		const strokes = [...ink.strokes];
		const anchor = structuredClone(ink.anchor);
		session.setInk(0, 0, anchor, { strokes: strokes.slice(1), ...size }, 'erase-1');
		session.setInk(0, 0, anchor, { strokes: strokes.slice(2), ...size }, 'erase-1');
		expect(inkOf(wb).strokes).toHaveLength(2);
		// The next gesture is a step of its own.
		session.setInk(0, 0, anchor, { strokes: strokes.slice(3), ...size }, 'erase-2');
		session.undo();
		expect(inkOf(wb).strokes).toHaveLength(2);
		session.undo();
		expect(inkOf(wb).strokes).toHaveLength(4);
		expect(session.canUndo()).toBe(false);
	});
});
