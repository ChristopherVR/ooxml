/**
 * A table made in code saves its column grid at its own frame's width, not at
 * a fixed 10 in. Each case is saved and loaded back, so the written
 * `a:gridCol` widths, the frame and the reloaded table agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { TablePptxElement } from '../../types';
import type { TableInput, TableOptions } from './types';

const EMU_PER_PX = 9525;

async function roundTrip(
	input: TableInput,
	options?: TableOptions,
): Promise<{ grid: number[]; frameEmu: number; table: TablePptxElement | undefined }> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addTable(input, options).build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const frame = /<p:graphicFrame>[\s\S]*?<a:ext cx="(\d+)"/u.exec(xml);
	const loaded = await new PptxHandler().load(bytes);
	return {
		grid: [...xml.matchAll(/<a:gridCol w="(\d+)"/gu)].map((m) => Number(m[1])),
		frameEmu: Number(frame?.[1]),
		table: loaded.slides[0]?.elements.find((e) => e.type === 'table') as
			| TablePptxElement
			| undefined,
	};
}

const sum = (values: number[]) => values.reduce((a, b) => a + b, 0);

describe('new table grid width', () => {
	it('sizes the columns from the frame width', async () => {
		// 400 px frame = 3 810 000 EMU; columns 100 / 300 px.
		const { grid, frameEmu, table } = await roundTrip(
			{ rows: [{ cells: [{ text: 'a' }, { text: 'b' }] }], columnWidths: [100, 300] },
			{ width: 400 },
		);
		expect(frameEmu).toBe(400 * EMU_PER_PX);
		expect(grid).toEqual([952500, 2857500]);
		expect(table?.width).toBe(400);
		expect(table?.tableData?.columnWidths).toEqual([0.25, 0.75]);
	});

	it('splits the frame evenly without widths', async () => {
		const { grid, table } = await roundTrip(
			{ rows: [{ cells: [{ text: 'a' }, { text: 'b' }, { text: 'c' }] }] },
			{ width: 300 },
		);
		expect(grid).toEqual([952500, 952500, 952500]);
		for (const fraction of table?.tableData?.columnWidths ?? []) {
			expect(fraction).toBeCloseTo(1 / 3, 6);
		}
	});

	it('uses the default frame width when none is given', async () => {
		const { grid, frameEmu } = await roundTrip({
			rows: [{ cells: [{ text: 'a' }, { text: 'b' }] }],
		});
		expect(sum(grid)).toBe(frameEmu);
		expect(grid[0]).toBe(grid[1]);
	});

	it('adds up to the frame when the fractions do not divide it evenly', async () => {
		const { grid, frameEmu } = await roundTrip(
			{ rows: [{ cells: Array.from({ length: 7 }, (_, i) => ({ text: String(i) })) }] },
			{ width: 101.5 },
		);
		expect(grid).toHaveLength(7);
		expect(Math.abs(sum(grid) - frameEmu)).toBeLessThanOrEqual(7);
		for (const width of grid) {
			expect(Math.abs(width - frameEmu / 7)).toBeLessThanOrEqual(1);
		}
	});

	it('writes one column at the full frame width', async () => {
		const { grid, frameEmu } = await roundTrip(
			{ rows: [{ cells: [{ text: 'only' }] }] },
			{ width: 250 },
		);
		expect(grid).toEqual([frameEmu]);
	});
});
