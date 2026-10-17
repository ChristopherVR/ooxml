import { beforeAll, describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { encodeBitmapEmf } from './emf-bitmap';
import type {
	VisioChartObjectInsertEdit,
	VisioChartObjectReplaceEdit,
} from './edit-chart-object-commands';
import {
	chartGridText,
	gridChartWorkbook,
	parseChartGrid,
	readEmbeddedChart,
} from '../xlsx/ui/dialogs/grid-chart';

let workbook: Uint8Array;
let other: Uint8Array;
const preview = encodeBitmapEmf(2, 2, new Uint8Array(16).fill(255));

beforeAll(async () => {
	workbook = (await gridChartWorkbook(parseChartGrid('Month\tSales\nJan\t10\nFeb\t20'), 'column', {
		title: 'First',
	}))!;
	other = (await gridChartWorkbook(parseChartGrid('Q\tRevenue\nA\t5\nB\t7\nC\t9'), 'line', {
		title: 'Second',
	}))!;
});

/** `preview: null` leaves the preview out; omitted, the test EMF is used. */
const insert = (
	patch: Omit<Partial<VisioChartObjectInsertEdit>, 'preview'> & {
		preview?: Uint8Array | null;
	} = {},
): VisioChartObjectInsertEdit => {
	const { preview: given = preview, ...rest } = patch;
	return {
		type: 'insert-chart-object',
		pageId: '0',
		shapeId: '1',
		x: 4.25,
		y: 5.5,
		width: 6.4,
		height: 4,
		workbook,
		...(given ? { preview: given } : {}),
		...rest,
	};
};
const replace = (
	patch: Partial<VisioChartObjectReplaceEdit> = {},
): VisioChartObjectReplaceEdit => ({
	type: 'replace-chart-object',
	pageId: '0',
	shapeId: '1',
	workbook: other,
	preview,
	...patch,
});
const text = async (bytes: Uint8Array, path: string) =>
	new TextDecoder().decode(await (await VisioPackage.open(bytes)).readBytes(path));
const hasPart = async (bytes: Uint8Array, path: string) => {
	try {
		await (await VisioPackage.open(bytes)).readBytes(path);
		return true;
	} catch {
		return false;
	}
};

describe('chart object insertion', () => {
	it('embeds the workbook and an EMF preview the way Visio stores Insert > Chart', async () => {
		const result = await editVsdx(await createVsdx(), [insert()]);
		expect(result.changedParts).toEqual(
			expect.arrayContaining([
				'visio/embeddings/Microsoft_Excel_Worksheet.xlsx',
				'visio/embeddings/_rels/Microsoft_Excel_Worksheet.xlsx.rels',
				'visio/media/image1.emf',
			]),
		);
		const types = await text(result.bytes, '[Content_Types].xml');
		expect(types).toContain('Extension="xlsx"');
		expect(types).toContain('<Default Extension="emf" ContentType="image/x-emf"/>');
		expect(await text(result.bytes, 'visio/pages/_rels/page1.xml.rels')).toMatch(
			/relationships\/package" Target="\.\.\/embeddings\/Microsoft_Excel_Worksheet\.xlsx"/,
		);
		expect(
			await text(result.bytes, 'visio/embeddings/_rels/Microsoft_Excel_Worksheet.xlsx.rels'),
		).toMatch(/relationships\/image" Target="\.\.\/media\/image1\.emf"/);
		const page = await text(result.bytes, 'visio/pages/page1.xml');
		expect(page).toContain('Type="Foreign"');
		expect(page).toMatch(/<ForeignData ForeignType="Object" ObjectType="33280" ShowAsIcon="0"/);
	});

	it('reads back as a chart the viewer can draw, with no unsupported-object warning', async () => {
		const { bytes } = await editVsdx(await createVsdx(), [insert()]);
		const document = await parseVsdx(bytes);
		const shape = document.pages[0]!.shapes[0]!;
		expect(shape).toMatchObject({ id: '1', kind: 'foreign', width: 6.4, height: 4 });
		expect(shape.embeddedObject).toMatchObject({ objectType: 33280, width: 6.4, height: 4 });
		expect(shape.embeddedObject!.chart.title).toBe('First');
		expect(shape.embeddedObject!.chart.series.map((s) => s.name)).toEqual(['Sales']);
		expect(document.diagnostics.map((d) => d.code)).not.toContain('unsupported-foreign-object');
		// The view model crosses the worker boundary, so it must be cloneable data.
		expect(() => structuredClone(shape.embeddedObject)).not.toThrow();
	});

	it('keeps the workbook bytes so the chart can be edited and put back', async () => {
		const { bytes } = await editVsdx(await createVsdx(), [insert()]);
		const shape = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		const chart = (await readEmbeddedChart(shape.embeddedObject!.bytes))!;
		expect(chartGridText(chart.rows)).toBe('Month\tSales\nJan\t10\nFeb\t20');
	});

	it('works without a preview, leaving no media part', async () => {
		const result = await editVsdx(await createVsdx(), [insert({ preview: null })]);
		expect(await hasPart(result.bytes, 'visio/media/image1.emf')).toBe(false);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.embeddedObject).toBeDefined();
	});

	it('chooses free part names and relationship IDs for a second chart', async () => {
		const first = (await editVsdx(await createVsdx(), [insert()])).bytes;
		const second = await editVsdx(first, [insert({ shapeId: '2', workbook: other })]);
		expect(second.changedParts).toContain('visio/embeddings/Microsoft_Excel_Worksheet1.xlsx');
		expect(second.changedParts).toContain('visio/media/image2.emf');
		const rels = await text(second.bytes, 'visio/pages/_rels/page1.xml.rels');
		expect(rels).toContain('Id="rId1"');
		expect(rels).toContain('Id="rId2"');
		const shapes = (await parseVsdx(second.bytes)).pages[0]!.shapes;
		expect(shapes.map((s) => s.embeddedObject?.chart.title)).toEqual(['First', 'Second']);
	});
});

describe('chart object replacement', () => {
	it('swaps the workbook and preview in place and keeps the placement', async () => {
		const inserted = (await editVsdx(await createVsdx(), [insert()])).bytes;
		const result = await editVsdx(inserted, [replace()]);
		expect(result.changedParts).toEqual(
			expect.arrayContaining([
				'visio/embeddings/Microsoft_Excel_Worksheet.xlsx',
				'visio/media/image1.emf',
			]),
		);
		// Nothing was added alongside the original parts.
		expect(await hasPart(result.bytes, 'visio/embeddings/Microsoft_Excel_Worksheet1.xlsx')).toBe(
			false,
		);
		expect(await hasPart(result.bytes, 'visio/media/image2.emf')).toBe(false);
		const shape = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect(shape).toMatchObject({ width: 6.4, height: 4 });
		expect(shape.embeddedObject!.chart.title).toBe('Second');
	});

	it('adds a preview when the object had none', async () => {
		const inserted = (await editVsdx(await createVsdx(), [insert({ preview: null })])).bytes;
		const result = await editVsdx(inserted, [replace()]);
		expect(result.changedParts).toContain('visio/media/image1.emf');
		expect(
			await text(result.bytes, 'visio/embeddings/_rels/Microsoft_Excel_Worksheet.xlsx.rels'),
		).toContain('image1.emf');
	});

	it('refuses a shape that is not an embedded chart', async () => {
		const inserted = (await editVsdx(await createVsdx(), [insert()])).bytes;
		await expect(editVsdx(inserted, [replace({ shapeId: '99' })])).rejects.toThrow(
			/not an embedded chart/,
		);
		await expect(editVsdx(inserted, [replace({ pageId: 'missing' })])).rejects.toThrow(/Page/);
	});
});

describe('chart object edits stay valid', () => {
	it('moves, resizes and deletes the chart with the existing edits', async () => {
		const inserted = (await editVsdx(await createVsdx(), [insert()])).bytes;
		const moved = await editVsdx(inserted, [
			{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 3 },
		]);
		const resized = await editVsdx(moved.bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 },
		]);
		const chart = (await parseVsdx(resized.bytes)).pages[0]!.shapes[0]!;
		expect(chart.rotation).toMatchObject({ pinX: 2, pinY: 3 });
		expect(chart.embeddedObject).toMatchObject({ width: 4, height: 2 });
		const removed = await editVsdx(resized.bytes, [
			{ type: 'delete-shape', pageId: '0', shapeId: '1' },
		]);
		expect((await parseVsdx(removed.bytes)).pages[0]!.shapes).toHaveLength(0);
	});

	it('refuses what it cannot safely embed', async () => {
		const base = await createVsdx();
		await expect(editVsdx(base, [insert({ workbook: new Uint8Array([1, 2, 3]) })])).rejects.toThrow(
			/Open XML package/,
		);
		await expect(editVsdx(base, [insert({ preview: new Uint8Array(64) })])).rejects.toThrow(/EMF/);
		await expect(editVsdx(base, [insert({ width: 0 })])).rejects.toThrow(/positive dimensions/);
		await expect(editVsdx(base, [insert({ pageId: 'nope' })])).rejects.toThrow(/Page/);
		await expect(
			editVsdx(base, [insert(), { type: 'move-shape', pageId: '0', shapeId: '1', x: 1, y: 1 }]),
		).rejects.toThrow(/own transaction/);
	});
});
