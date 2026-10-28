import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { chartSpaceDataSheet, chartExSpaceDataSheet } from 'ooxml-core/chart';
import {
	readEmbeddedChartSpace,
	readEmbeddedChartExSpace,
	loadXlsx,
	saveXlsx,
} from 'ooxml-core/xlsx';
import { gridChartWorkbook, readEmbeddedChart } from 'ooxml-core/xlsx/ui';
import { registerOfficeUi } from '../index';
import { ChartObjectEditor } from './viewer-chart-editor';
import type { ChartObjectData, ChartTarget, RasterizeSvg } from './viewer-insert-chart';

afterEach(() => document.body.replaceChildren());

const raster: RasterizeSvg = async () => ({
	width: 2,
	height: 2,
	rgba: new Uint8ClampedArray(16).fill(255),
});

type Submit = (data: ChartObjectData, target: ChartTarget) => Promise<void>;
const TARGET: ChartTarget = { pageId: '0', shapeId: '7', width: 4, height: 2.5, generation: 1 };
const ROWS = [
	[null, 'Plan', 'Actual'],
	['North', 4, 2],
	['South', 3, 5],
];

function setup(submit: Submit = vi.fn(async () => {})) {
	registerOfficeUi();
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const editor = new ChartObjectEditor(root, submit, raster);
	const press = (label: string) =>
		editor.dialog
			.querySelector(`[command="chart-edit-${label.toLowerCase()}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	return { editor, submit, press };
}

const isEmf = (bytes: Uint8Array) =>
	new DataView(bytes.buffer, bytes.byteOffset).getUint32(40, true) === 0x464d4520;

describe('chart object Edit Data and Format Chart', () => {
	it('edits embedded ChartEx through the shared data sheet and formatting pane', async () => {
		const book = await loadXlsx(
			readFileSync('../core/xlsx/__fixtures__/excel-chart-families.xlsx'),
		);
		for (const ws of book.sheets)
			ws.drawings = ws.drawings.filter(
				(d) =>
					d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
			);
		const submit = vi.fn<Submit>(async () => {});
		const { editor, press } = setup(submit);
		expect(await editor.open('data', TARGET, await saveXlsx(book))).toBe(true);
		expect(editor.space).toBeUndefined();
		expect(editor.chartExSpace?.plotArea.series[0]?.layoutId).toBe('waterfall');
		editor.sheet.dispatchEvent(
			new CustomEvent('office-chart-data-change', {
				detail: { edits: [{ kind: 'value', series: 0, category: 0, value: 42 }] },
			}),
		);
		editor.pane.dispatchEvent(
			new CustomEvent('office-chart-format-edit', {
				detail: {
					intent: {
						op: 'fill',
						target: { kind: 'series', seriesIndex: 0 },
						fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
					},
				},
			}),
		);
		expect(editor.preview.querySelector('svg')).not.toBeNull();
		press('OK');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		const back = (await readEmbeddedChartExSpace(submit.mock.calls[0]![0].workbook))!;
		expect(chartExSpaceDataSheet(back.space).series[0]!.values[0]).toBe(42);
		expect(isEmf(submit.mock.calls[0]![0].preview!)).toBe(true);
	});
	it('saves advanced shared chart settings without changing the embedded data', async () => {
		const submit = vi.fn<Submit>(async () => {});
		const { editor, press } = setup(submit);
		await editor.open('format', TARGET, (await gridChartWorkbook(ROWS, 'column'))!);
		for (const intent of [
			{ op: 'trendline', seriesIndex: 0, index: 0, edit: { type: 'linear' } },
			{
				op: 'errorBars',
				seriesIndex: 0,
				direction: 'y',
				edit: { valueType: 'fixedVal', value: 2 },
			},
		] satisfies import('ooxml-core/chart').ChartEditIntent[])
			editor.pane.dispatchEvent(
				new CustomEvent('office-chart-format-edit', {
					detail: { target: 'series:0', field: 'trendline.type', value: '', intent },
				}),
			);
		press('OK');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		const back = (await readEmbeddedChartSpace(submit.mock.calls[0]![0].workbook))!;
		expect(back.space.plotArea.groups[0]!.series[0]!.trendlines?.[0]?.trendlineType).toBe('linear');
		expect(back.space.plotArea.groups[0]!.series[0]!.errorBars?.[0]?.value).toBe(2);
		expect(chartSpaceDataSheet(back.space).series[0]?.values).toEqual([4, 3]);
	});
	it('edits the data in the shared data sheet and re-embeds the workbook', async () => {
		const submit = vi.fn<Submit>(async () => {});
		const { editor, press } = setup(submit);
		const workbook = (await gridChartWorkbook(ROWS, 'column'))!;
		expect(await editor.open('data', TARGET, workbook)).toBe(true);
		expect(editor.dialog.getAttribute('heading')).toBe('Edit Data');
		expect(editor.sheet.hidden).toBe(false);
		expect(editor.pane.hidden).toBe(true);
		expect(editor.sheet.data.series.map((series) => series.name)).toEqual(['Plan', 'Actual']);
		expect(editor.preview.querySelector('svg')).not.toBeNull();
		editor.sheet.dispatchEvent(
			new CustomEvent('office-chart-data-change', {
				detail: {
					edits: [{ kind: 'value', series: 0, category: 1, value: 8 }, { kind: 'add-category' }],
					data: editor.sheet.data,
				},
			}),
		);
		expect(editor.sheet.data.categories).toEqual(['North', 'South', 'Cat 3']);
		press('OK');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		const [data, target] = submit.mock.calls[0]!;
		expect(target).toBe(TARGET);
		expect(isEmf(data.preview)).toBe(true);
		const embedded = (await readEmbeddedChart(data.workbook))!;
		expect(embedded.rows[2]).toEqual(['South', 8, 5]);
		expect(embedded.view.series[0]?.values).toEqual([4, 8, 0]);
		expect(editor.dialog.open).toBe(false);
	});

	it('formats the chart through the shared Format pane intents', async () => {
		const submit = vi.fn<Submit>(async () => {});
		const { editor, press } = setup(submit);
		const workbook = (await gridChartWorkbook(ROWS, 'column'))!;
		await editor.open('format', TARGET, workbook);
		expect(editor.dialog.getAttribute('heading')).toBe('Format Chart');
		expect(editor.pane.hidden).toBe(false);
		editor.pane.dispatchEvent(
			new CustomEvent('office-chart-format-edit', {
				detail: {
					target: 'series:1',
					field: 'fill.color',
					value: '#FF0000',
					intent: {
						op: 'fill',
						target: { kind: 'series', seriesIndex: 1 },
						fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
					},
				},
			}),
		);
		editor.pane.dispatchEvent(
			new CustomEvent('office-chart-format-edit', {
				detail: {
					target: 'legend',
					field: 'legend.position',
					value: 't',
					intent: { op: 'legend', edit: { position: 't' } },
				},
			}),
		);
		expect(editor.space?.legend?.position).toBe('t');
		press('OK');
		await vi.waitFor(() => expect(submit).toHaveBeenCalled());
		const back = (await readEmbeddedChartSpace(submit.mock.calls[0]![0].workbook))!;
		expect(back.space.plotArea.groups[0]!.series[1]!.spPr?.fill).toMatchObject({
			kind: 'solid',
			color: { value: 'FF0000' },
		});
		expect(back.space.legend?.position).toBe('t');
		expect(chartSpaceDataSheet(back.space).series[0]?.values).toEqual([4, 3]);
	});

	it('refuses a workbook with no chart and shows a failed save', async () => {
		const submit = vi.fn<Submit>(async () => {
			throw new Error('The drawing changed.');
		});
		const { editor, press } = setup(submit);
		expect(await editor.open('data', TARGET, new Uint8Array([1, 2]))).toBe(false);
		await editor.open('data', TARGET, (await gridChartWorkbook(ROWS, 'column'))!);
		press('OK');
		await vi.waitFor(() => expect(editor.error.textContent).toBe('The drawing changed.'));
		expect(editor.dialog.open).toBe(true);
		press('Cancel');
		expect(editor.dialog.open).toBe(false);
	});
});

it('shows unsupported chart content as read-only and allows a later supported chart', async () => {
	const submit = vi.fn<Submit>(async () => {});
	const { editor, press } = setup(submit);
	const ribbon = document.createElement('office-ui-ribbon');
	ribbon.dataset.visioLocale = 'fr';
	(editor.dialog.parentNode as ShadowRoot).append(ribbon);
	const workbook = (await gridChartWorkbook(ROWS, 'column'))!;
	const zip = await JSZip.loadAsync(workbook);
	const part = Object.keys(zip.files).find((name) => /^xl\/charts\/chart\d+\.xml$/.test(name))!;
	zip.file(
		part,
		(await zip.file(part)!.async('string')).replace('</c:ser>', '<c:unknown/></c:ser>'),
	);
	const unsupported = await zip.generateAsync({ type: 'uint8array' });
	expect(await editor.open('format', TARGET, unsupported)).toBe(false);
	expect(editor.dialog.open).toBe(true);
	expect(editor.error.textContent).toContain('lecture seule');
	expect(editor.sheet.hidden).toBe(true);
	expect(editor.pane.hidden).toBe(true);
	expect(editor.space).toBeUndefined();
	press('OK');
	expect(submit).not.toHaveBeenCalled();
	press('Cancel');
	expect(await editor.open('format', TARGET, workbook)).toBe(true);
	expect(editor.error.textContent).toBe('');
	press('OK');
	await vi.waitFor(() => expect(submit).toHaveBeenCalledOnce());
});
