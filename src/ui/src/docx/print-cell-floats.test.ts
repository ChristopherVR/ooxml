// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer, layoutDocumentModel } from 'ooxml-core/docx/layout';
import {
	CHART_PARTS,
	cellFloatDemo,
	cellFloatPackage,
} from '../../../core/docx/__fixtures__/cell-floats/cases';
import { renderPrintLayout } from './print-layout';

const px = (value: string) => Number.parseFloat(value) || 0;

async function demo() {
	const chart = await JSZip.loadAsync(
		readFileSync(join(process.cwd(), '../core/docx/__fixtures__/chart.docx')),
	);
	const parts: Record<string, string> = {};
	for (const part of CHART_PARTS) parts[part] = await chart.file(part)!.async('string');
	const zip = new JSZip();
	for (const [path, xml] of Object.entries(cellFloatPackage(cellFloatDemo(), parts)))
		zip.file(path, xml);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const result = layoutDocumentModel(model, createFakeMeasurer());
	return { model, result, ...renderPrintLayout(result, undefined, {}) };
}

describe('Print Layout floating objects in table cells', () => {
	it('draws objects anchored in cells over their cells, and floating charts as charts', async () => {
		const { element, renderAll } = await demo();
		renderAll();
		const sheet = element.querySelector<HTMLElement>('.dve-print-page')!;
		const floats = [...sheet.querySelectorAll<HTMLElement>(':scope > .dve-print-float')];
		// A square shape, a chart, a top-and-bottom shape and a shape behind text in cells, and a
		// chart floating in the body.
		expect(floats).toHaveLength(5);
		const charts = floats.filter((float) => float.querySelector('svg') && float.title !== 'Shape');
		expect(charts.length).toBeGreaterThanOrEqual(2);
		expect(floats[3]!.classList.contains('dve-print-float-behind')).toBe(true);

		const column = sheet.querySelector<HTMLElement>('.dve-print-column')!;
		const tables = [...column.querySelectorAll<HTMLElement>(':scope > .dve-print-table')];
		expect(tables).toHaveLength(4);
		for (const [index, table] of tables.entries()) {
			const float = floats[index]!;
			const row = table.querySelector<HTMLElement>('.dve-print-row')!;
			const cell = row.querySelectorAll<HTMLElement>('.dve-print-cell')[1]!;
			const cellLeft = px(column.style.left) + px(table.style.left) + px(cell.style.left);
			const rowTop = px(column.style.top) + px(table.style.top) + px(row.style.top);
			const left = px(float.style.left);
			const top = px(float.style.top);
			// Inside the anchor cell, and the cell's row is at least as tall as a wrapping object.
			expect(left, `float ${index} left`).toBeGreaterThanOrEqual(cellLeft - 0.5);
			expect(left + px(float.style.width)).toBeLessThanOrEqual(
				cellLeft + px(cell.style.width) + 0.5,
			);
			expect(top, `float ${index} top`).toBeGreaterThanOrEqual(rowTop - 0.5);
			if (index < 3)
				expect(top + px(float.style.height)).toBeLessThanOrEqual(
					rowTop + px(row.style.height) + 0.5,
				);
		}
	});

	it('wraps the cell text beside a square-wrapped object', async () => {
		const { result } = await demo();
		const page = result.pages[0]!;
		const square = page.floats![0]!;
		const table = page.columns[0]!.blocks.find((block) => block.kind === 'table')!;
		if (table.kind !== 'table') throw new Error('Expected the first table');
		const row = table.rows[0]!;
		const geometry = row.geometry![1]!;
		const anchor = row.cells[1]!.at(-1)!;
		const contentLeft =
			page.marginLeftPx +
			page.columns[0]!.xPx +
			(table.xPx ?? 0) +
			geometry.xPx +
			geometry.paddingLeftPx;
		const top = page.marginTopPx + table.yPx + row.yPx + anchor.yPx;
		const beside = anchor.lines.filter(
			(line) =>
				top + line.yPx < square.yPx + square.heightPx &&
				top + line.yPx + line.heightPx > square.yPx,
		);
		expect(beside.length).toBeGreaterThan(0);
		// Each line beside the object starts right of it (and its 9 pt distance).
		for (const line of beside)
			expect(contentLeft + line.fragments[0]!.xPx).toBeGreaterThanOrEqual(
				square.xPx + square.widthPx + 12 - 0.5,
			);
	});
});
