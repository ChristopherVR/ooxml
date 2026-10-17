import { afterEach, describe, expect, it, vi } from 'vitest';
import { gridChartWorkbook, parseChartGrid, readEmbeddedChart } from 'ooxml-core/xlsx/ui';
import { registerOfficeUi } from '../index';
import {
	InsertChartDialog,
	type ChartObjectData,
	type ChartTarget,
	type RasterizeSvg,
} from './viewer-insert-chart';

afterEach(() => document.body.replaceChildren());

const raster: RasterizeSvg = async () => ({
	width: 2,
	height: 2,
	rgba: new Uint8ClampedArray(16).fill(255),
});

type Submit = (data: ChartObjectData, target: ChartTarget | undefined) => Promise<void>;

function setup(submit: Submit = vi.fn(async () => {}), rasterize: RasterizeSvg = raster) {
	registerOfficeUi();
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const drawn = vi.fn(rasterize);
	const chart = new InsertChartDialog(root, submit, drawn);
	const press = (label: string) =>
		chart.dialog
			.querySelector(`[command="chart-dialog-${label.toLowerCase()}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	const edit = (data: string) => {
		chart.picker.value = { ...chart.picker.value, data };
	};
	return { chart, submit, drawn, press, edit };
}

const isEmf = (bytes: Uint8Array) =>
	new DataView(bytes.buffer, bytes.byteOffset).getUint32(40, true) === 0x464d4520;

describe('Insert > Chart dialog', () => {
	it('opens on sample data with a live preview for each type', () => {
		const { chart } = setup();
		chart.open();
		expect(chart.dialog.open).toBe(true);
		expect(chart.dialog.getAttribute('heading')).toBe('Insert Chart');
		expect(chart.picker.value.data).toContain('Series 1');
		expect(chart.picker.shadowRoot!.querySelector('.preview svg')).not.toBeNull();
		expect(chart.picker.shadowRoot!.querySelectorAll('[data-gallery-item]')).toHaveLength(8);
	});

	it('hands over an Excel workbook that holds the chart, and an EMF preview', async () => {
		const submit = vi.fn<Submit>(async () => {});
		const { chart, drawn, press } = setup(submit);
		chart.open();
		chart.picker.value = { ...chart.picker.value, type: 'line', title: 'Revenue' };
		press('Insert');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		const [data, target] = submit.mock.calls[0]!;
		expect(target).toBeUndefined();
		expect(isEmf(data.preview)).toBe(true);
		const embedded = (await readEmbeddedChart(data.workbook))!;
		expect(embedded.chart).toMatchObject({ chartType: 'line', title: 'Revenue' });
		expect(embedded.rows[0]).toEqual(['Category', 'Series 1', 'Series 2']);
		const [, svg, width, height] = drawn.mock.calls[0]!;
		expect(svg.startsWith('<svg')).toBe(true);
		expect([width, height]).toEqual([640, 400]);
		await vi.waitFor(() => expect(chart.dialog.open).toBe(false));
	});

	it('asks for data instead of inserting an empty chart', () => {
		const { chart, submit, press, edit } = setup();
		chart.open();
		edit('   ');
		press('Insert');
		expect(chart.error.textContent).toMatch(/header row/);
		edit('only a header');
		press('Insert');
		expect(chart.error.textContent).toMatch(/header row/);
		expect(submit).not.toHaveBeenCalled();
		expect(chart.dialog.open).toBe(true);
	});

	it('keeps the dialog open and says why when the editor refuses the chart', async () => {
		const submit = vi.fn<Submit>(async () => {
			throw new Error('The chart is larger than the 8 MiB embedding limit.');
		});
		const { chart, press } = setup(submit);
		chart.open();
		press('Insert');
		await vi.waitFor(() => expect(chart.error.textContent).toMatch(/8 MiB/));
		expect(chart.dialog.open).toBe(true);
	});

	it('reports a browser that cannot draw the chart', async () => {
		const failing: RasterizeSvg = async () => {
			throw new Error('This browser cannot draw the chart.');
		};
		const { chart, submit, press } = setup(undefined, failing);
		chart.open();
		press('Insert');
		await vi.waitFor(() => expect(chart.error.textContent).toMatch(/cannot draw/));
		expect(submit).not.toHaveBeenCalled();
	});

	it('Cancel closes without inserting', () => {
		const { chart, submit, press } = setup();
		chart.open();
		press('Cancel');
		expect(chart.dialog.open).toBe(false);
		expect(submit).not.toHaveBeenCalled();
	});
});

describe('Edit Chart', () => {
	const target: ChartTarget = { pageId: '0', shapeId: '7', width: 6, height: 3, generation: 1 };
	const existing = async () =>
		(await gridChartWorkbook(parseChartGrid('Month\tSales\nJan\t10\nFeb\t20'), 'bar', {
			title: 'Kept',
		}))!;

	it("opens with the chart's own type, title and data, and says what saving does", async () => {
		const { chart } = setup();
		expect(await chart.edit(target, await existing())).toBe(true);
		expect(chart.dialog.getAttribute('heading')).toBe('Edit Chart');
		expect(chart.picker.value).toMatchObject({
			type: 'bar',
			title: 'Kept',
			data: 'Month\tSales\nJan\t10\nFeb\t20',
		});
		expect(chart.dialog.querySelector('.chart-note')!.textContent).toMatch(/not kept/);
		expect(chart.target).toEqual(target);
	});

	it('submits against the same shape with a preview in the shape proportions', async () => {
		const submit = vi.fn<Submit>(async () => {});
		const { chart, drawn, press, edit } = setup(submit);
		await chart.edit(target, await existing());
		edit('Month\tSales\nJan\t30\nFeb\t40\nMar\t50');
		// The confirm button keeps its command but is relabelled OK while editing.
		expect(
			chart.dialog.querySelector('[command="chart-dialog-insert"]')!.getAttribute('label'),
		).toBe('OK');
		press('Insert');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		expect(submit.mock.calls[0]![1]).toEqual(target);
		// 6 x 3 inches is 2:1, so a 640 pixel wide preview is 320 high.
		expect(drawn.mock.calls[0]!.slice(2)).toEqual([640, 320]);
		const embedded = (await readEmbeddedChart(submit.mock.calls[0]![0].workbook))!;
		expect(embedded.rows).toHaveLength(4);
	});

	it('refuses an object whose workbook has no chart', async () => {
		const { chart } = setup();
		expect(await chart.edit(target, new Uint8Array([1, 2, 3]))).toBe(false);
		expect(chart.dialog.open).toBe(false);
	});

	it('a new chart after an edit starts clean', async () => {
		const { chart } = setup();
		await chart.edit(target, await existing());
		chart.close();
		chart.open();
		expect(chart.dialog.getAttribute('heading')).toBe('Insert Chart');
		expect(chart.target).toBeUndefined();
		expect(chart.dialog.querySelector<HTMLElement>('.chart-note')!.hidden).toBe(true);
	});
});
