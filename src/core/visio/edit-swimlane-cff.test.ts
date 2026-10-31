import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import type { VisioShape } from './model';
import { parseXml } from '../xml/index';
import { attribute, children } from './sheet';
import { fixture } from './test-fixtures';

/**
 * The oracle files are what Visio 16 saved after each change to the Cross-Functional Flowchart
 * template chart, made through COM with events enabled (see the .md next to the fixtures): a
 * Separator dropped at x = 8 in, the title bar off, the phase headings off, style 3, rotated
 * lane labels.
 */
const bytesOf = (name: string) =>
	new Uint8Array(readFileSync(new URL(`./__fixtures__/${name}.vsdx`, import.meta.url)));
const template = () => bytesOf('cross-functional-native');

/** Every cell of a page, by `shape/sub-shape section row cell`, as Visio saved it (V attribute). */
async function cellsOf(bytes: Uint8Array): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const root = parseXml(await zip.file('visio/pages/page1.xml')!.async('string')).documentElement;
	const out = new Map<string, string>();
	const walk = (shape: Element, prefix: string) => {
		const key = prefix + attribute(shape, 'ID');
		const put = (name: string, cell: Element) => out.set(name, cell.getAttribute('V') ?? '');
		for (const cell of children(shape, 'Cell')) put(`${key} ${attribute(cell, 'N')}`, cell);
		for (const section of children(shape, 'Section')) {
			const name = `${attribute(section, 'N')}${attribute(section, 'IX') ?? ''}`;
			for (const cell of children(section, 'Cell'))
				put(`${key} ${name} ${attribute(cell, 'N')}`, cell);
			for (const row of children(section, 'Row')) {
				const rowName = attribute(row, 'N') ?? `IX${attribute(row, 'IX')}`;
				for (const cell of children(row, 'Cell'))
					put(`${key} ${name} ${rowName} ${attribute(cell, 'N')}`, cell);
			}
		}
		for (const sub of children(children(shape, 'Shapes')[0], 'Shape')) walk(sub, `${key}/`);
	};
	for (const shape of children(children(root, 'Shapes')[0], 'Shape')) walk(shape, '');
	return out;
}

const same = (a: string, b: string | undefined) => {
	if (b === undefined) return Number(a) === 0 && a !== '';
	const x = Number(a);
	const y = Number(b);
	return Number.isFinite(x) && Number.isFinite(y) && a !== '' && b !== ''
		? Math.abs(x - y) <= 1e-6
		: a === b;
};

/** The cells Visio changed that our edit of the same file did not change the same way. */
async function mismatches(ours: Uint8Array, visio: Uint8Array): Promise<string[]> {
	const [mine, before, theirs] = await Promise.all([
		cellsOf(ours),
		cellsOf(template()),
		cellsOf(visio),
	]);
	const out: string[] = [];
	for (const [key, value] of theirs) {
		const was = before.get(key);
		if (was !== undefined && same(value, was)) continue;
		if (!same(value, mine.get(key))) out.push(`${key}: Visio ${value}, ours ${mine.get(key)}`);
	}
	return out;
}

const run = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
const options = (change: object): VisioEdit =>
	({ type: 'set-swimlane-options', pageId: '0', frameId: '1', ...change }) as VisioEdit;
const separator = (shapeId: string, extra: object = {}): VisioEdit =>
	({ type: 'insert-separator', pageId: '0', frameId: '1', shapeId, ...extra }) as VisioEdit;

async function shapesOf(bytes: Uint8Array): Promise<Map<string, VisioShape>> {
	return new Map((await parseVsdx(bytes)).pages[0]!.shapes.map((shape) => [shape.id, shape]));
}
const box = (shape: VisioShape) => {
	const [a, , , d, e, f] = shape.transform;
	return { minX: e, maxX: e + a * shape.width, minY: f, maxY: f + d * shape.height };
};
const pageXml = async (bytes: Uint8Array) =>
	await (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');

describe('Insert Phase on a chart Visio built', () => {
	it('splits the phase like Visio: the new separator takes the part up to the drop point', async () => {
		const ours = await run(template(), [separator('15', { at: 8 })]);
		// Everything Visio changed matches to the last cell (widths, pins, sub-shape geometry,
		// Relationships); the one difference is the fill the drop took from the theme variant.
		expect(await mismatches(ours, bytesOf('cross-functional-native-phase'))).toEqual([
			'15 FillBkgnd: Visio #4eb4e2, ours #4d4d4d',
			'15/16 FillBkgnd: Visio #4eb4e2, ours #4d4d4d',
			'15/17 FillBkgnd: Visio #4eb4e2, ours #4d4d4d',
		]);
		const map = await shapesOf(ours);
		expect(map.get('5')!.structure).toMatchObject({ type: 'list', memberIds: ['15', '12'] });
		expect(box(map.get('15')!).maxX).toBeCloseTo(8);
		expect(box(map.get('12')!).minX).toBeCloseTo(8);
		const xml = await pageXml(ours);
		expect(xml).toContain('DEPENDSON(2,Sheet.15!SheetRef(),Sheet.12!SheetRef())');
		expect(map.get('15')!.masterId).toBe(map.get('12')!.masterId);
	});

	it('moves the first phase heading extension to the new first phase', async () => {
		const first = await run(template(), [separator('15', { at: 8 })]);
		const before = await cellsOf(template());
		expect(before.get('12/14 User FirstHeadingExtend Value')).toBeDefined();
		const after = await cellsOf(first);
		expect(after.has('12/14 User FirstHeadingExtend Value')).toBe(false);
		expect(Number(after.get('15/17 User FirstHeadingExtend Value'))).toBeCloseTo(0.3937, 3);
		// A phase inserted into the second one leaves the extension where it is.
		const third = await cellsOf(await run(first, [separator('18', { at: 10 })]));
		expect(third.has('18/20 User FirstHeadingExtend Value')).toBe(false);
		expect(third.has('15/17 User FirstHeadingExtend Value')).toBe(true);
	});

	it('lists the separators in phase order and keeps them edge to edge', async () => {
		let bytes = await run(template(), [separator('15', { at: 8 })]);
		bytes = await run(bytes, [separator('18', { at: 3 }), separator('21', { at: 10 })]);
		const map = await shapesOf(bytes);
		const list = map.get('5')!;
		expect(list.structure).toMatchObject({ memberIds: ['18', '15', '21', '12'] });
		const boxes = ['18', '15', '21', '12'].map((id) => box(map.get(id)!));
		const span = box(list);
		expect(boxes[0]!.minX).toBeCloseTo(span.minX);
		for (let i = 1; i < boxes.length; i++) expect(boxes[i]!.minX).toBeCloseTo(boxes[i - 1]!.maxX);
		expect(boxes[3]!.maxX).toBeCloseTo(span.maxX);
		expect(boxes.map((b) => b.maxX - b.minX).every((w) => w >= 0.4 - 1e-9)).toBe(true);
	});

	it('splits the widest phase in the middle without a drop point, and names a phase', async () => {
		const bytes = await run(template(), [separator('15', { label: 'Build' })]);
		const map = await shapesOf(bytes);
		const kept = box(map.get('12')!);
		const made = box(map.get('15')!);
		expect(made.maxX).toBeCloseTo(kept.minX);
		expect(made.maxX - made.minX).toBeCloseTo(kept.maxX - kept.minX);
		expect(await pageXml(bytes)).toContain('>Build</Text>');
	});

	it('keeps a margin: a drop point near a boundary is moved in, a narrow phase is refused', async () => {
		const bytes = await run(template(), [separator('15', { at: 0.9 })]);
		const map = await shapesOf(bytes);
		expect(box(map.get('15')!).maxX - box(map.get('15')!).minX).toBeCloseTo(0.4);
		let narrow: Uint8Array = template();
		for (let i = 0; i < 40; i++) {
			try {
				narrow = await run(narrow, [separator(String(15 + i * 3), { at: 1 })]);
			} catch (error) {
				expect(String(error)).toMatch(/wide enough/);
				return;
			}
		}
		throw new Error('the phases never ran out of room');
	});
});

describe('chart options on a chart Visio built', () => {
	// Cells that hold a measured text height or an arc of a heading drawn from it: the sheet's
	// text measure is not available here, so the cache Visio recomputed is not reproduced.
	it('hides the title bar: the frame loses the bar at the top, like Visio', async () => {
		const ours = await run(template(), [options({ showTitle: false })]);
		expect(await mismatches(ours, bytesOf('cross-functional-native-notitle'))).toEqual([
			'1/3 User visHeadingHeight Value: Visio 0.3778313381618924, ours undefined',
			'1/3 Geometry1 IX2 D: Visio 1, ours 13.8499999994481',
		]);
		const shape = (await shapesOf(ours)).get('1')!;
		expect(shape.structure).toMatchObject({ swimlane: { options: { showTitle: false } } });
		const back = await run(ours, [options({ showTitle: true })]);
		expect(shape.height).toBeLessThan((await shapesOf(back)).get('1')!.height);
		expect((await shapesOf(back)).get('1')!.height).toBeCloseTo(
			(await shapesOf(template())).get('1')!.height,
		);
	});

	it('hides the phase headings: lanes rise under the frame top, the frame and Phase List shrink', async () => {
		const ours = await run(template(), [options({ showSeparators: false })]);
		expect(await mismatches(ours, bytesOf('cross-functional-native-nophases'))).toEqual([]);
		const back = await run(ours, [options({ showSeparators: true })]);
		const original = await shapesOf(template());
		const restored = await shapesOf(back);
		for (const id of ['1', '4', '5', '6', '9', '12'])
			expect(box(restored.get(id)!)).toEqual(
				expect.objectContaining({
					minY: expect.closeTo(box(original.get(id)!).minY, 6),
					maxY: expect.closeTo(box(original.get(id)!).maxY, 6),
				}),
			);
	});

	it('applies a style: the frame, lists, lanes and separators follow through their formulas', async () => {
		const ours = await run(template(), [options({ style: 3 })]);
		expect(await mismatches(ours, bytesOf('cross-functional-native-style3'))).toEqual([]);
		expect((await shapesOf(ours)).get('1')!.structure).toMatchObject({
			swimlane: { options: { style: 3 } },
		});
		expect(await pageXml(ours)).toContain('/stg1=3');
	});

	it('rotates the lane labels through the Swimlane List', async () => {
		const ours = await run(template(), [options({ rotateLabels: true })]);
		expect(await mismatches(ours, bytesOf('cross-functional-native-rotated'))).toEqual([
			'6/8 User visHeadingHeight Value: Visio 0.7111656887478298, ours undefined',
			'9/11 User visHeadingHeight Value: Visio 0.7111656887478298, ours undefined',
		]);
		expect(await pageXml(ours)).toContain('/stg6=1');
		// Visio keeps the flag on the Swimlane List, and so does the model.
		expect((await shapesOf(ours)).get('4')!.structure).toMatchObject({ rotateLabels: true });
	});

	it('takes several options at once and leaves unchanged ones alone', async () => {
		const ours = await run(template(), [
			options({ style: 2, showTitle: false, showSeparators: false, rotateLabels: true }),
		]);
		const info = (await shapesOf(ours)).get('1')!.structure;
		expect(info).toMatchObject({
			swimlane: { options: { style: 2, showTitle: false, showSeparators: false } },
		});
		const same = await run(template(), [options({ showTitle: true, showSeparators: true })]);
		expect(await mismatches(same, template())).toEqual([]);
	});

	it('refuses Direction: Visio swaps in vertical masters this editor does not copy', async () => {
		await expect(editVsdx(template(), [options({ direction: 'vertical' })])).rejects.toThrow(
			/Direction of a chart Visio built/,
		);
		// Asking for the direction the chart already has is not a change.
		await expect(
			editVsdx(template(), [options({ direction: 'horizontal' })]),
		).resolves.toBeDefined();
	});
});

describe('lane edits keep a chart Visio built laid out as Visio does', () => {
	it('keeps the phase headings above the lanes and the Phase List after the lane headings', async () => {
		const original = await shapesOf(template());
		const bytes = await run(template(), [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '1',
				shapeId: '30',
				relativeTo: '6',
				position: 'after',
				heading: 'New',
			},
		]);
		const map = await shapesOf(bytes);
		const lanesTop = Math.max(...['6', '9', '30'].map((id) => box(map.get(id)!).maxY));
		expect(box(map.get('12')!).maxY).toBeCloseTo(lanesTop + 5 / 25.4);
		expect(box(map.get('5')!).minX).toBeCloseTo(box(original.get('5')!).minX);
		expect(box(map.get('5')!).maxY).toBeCloseTo(box(map.get('12')!).maxY);
		expect(box(map.get('12')!).minY).toBeCloseTo(box(map.get('30')!).minY);
	});

	it('inserts a lane with the phase headings hidden: no band above the lanes', async () => {
		const hidden = await run(template(), [options({ showSeparators: false })]);
		const bytes = await run(hidden, [
			{
				type: 'insert-swimlane',
				pageId: '0',
				frameId: '1',
				shapeId: '30',
				relativeTo: '6',
				position: 'after',
				heading: 'New',
			},
		]);
		const map = await shapesOf(bytes);
		const lanesTop = Math.max(...['6', '9', '30'].map((id) => box(map.get(id)!).maxY));
		expect(box(map.get('12')!).maxY).toBeCloseTo(lanesTop);
		expect(box(map.get('5')!).maxY).toBeCloseTo(lanesTop);
	});
});

describe('phases and options on a chart the editor created', () => {
	const rect = (shapeId: string, x: number, y: number): VisioEdit => ({
		type: 'create-rectangle',
		pageId: '0',
		shapeId,
		x,
		y,
		width: 1,
		height: 0.6,
	});
	async function created(): Promise<Uint8Array> {
		const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
		const start = await run(blank, [rect('1', 4, 6.8)]);
		return run(start, [
			{
				type: 'create-swimlanes',
				pageId: '0',
				frameId: '50',
				laneIds: ['51', '52'],
				headings: ['A', 'B'],
				box: { x: 6, y: 5, width: 8, height: 4.35 },
				direction: 'horizontal',
				title: 'Flow',
				listId: '54',
				phaseListId: '55',
			},
		]);
	}
	const frameOptions = (change: object): VisioEdit =>
		({ type: 'set-swimlane-options', pageId: '0', frameId: '50', ...change }) as VisioEdit;
	const bar = (shapeId: string, at: number, label: string): VisioEdit =>
		({ type: 'insert-separator', pageId: '0', frameId: '50', shapeId, at, label }) as VisioEdit;

	it('writes Visio flowchart cells as plain numbers and keeps them readable', async () => {
		const bytes = await run(await created(), [
			frameOptions({ style: 3, showTitle: false, showSeparators: false, rotateLabels: true }),
		]);
		const cells = await cellsOf(bytes);
		expect(cells.get('50 User visCFFStyle Value')).toBe('3');
		expect(cells.get('50 User visShowTitle Value')).toBe('0');
		expect(cells.get('50 User visShowPhase Value')).toBe('0');
		// The Swimlane List holds the label flag, as Visio's does.
		expect(cells.get('54 User visRotateLabel Value')).toBe('1');
		const xml = await pageXml(bytes);
		expect(xml).not.toMatch(/N="visShowPhase"><Cell N="Value" V="0" U=/);
		expect((await shapesOf(bytes)).get('50')!.structure).toMatchObject({
			swimlane: {
				options: { style: 3, showTitle: false, showSeparators: false, rotateLabels: true },
			},
		});
	});

	it('inserts phases into the Phase List in order, naming them', async () => {
		const bytes = await run(await created(), [bar('60', 8, 'Plan'), bar('61', 4, 'Build')]);
		const map = await shapesOf(bytes);
		expect(map.get('55')!.structure).toMatchObject({ type: 'list', memberIds: ['61', '60'] });
		const xml = await pageXml(bytes);
		expect(xml).toContain('Plan');
		expect(xml).toContain('Build');
		expect(map.get('60')!.structure).toMatchObject({ swimlane: { role: 'separator' } });
	});
});
