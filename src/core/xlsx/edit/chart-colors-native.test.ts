import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-palette-edits.json';
import { chartView } from '../layout/chart-view';
import { THEME_SLOTS } from '../layout/colors';
import type { ChartObject, ChartType } from '../model';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { chartColorPalettePatch } from './chart-colors';
import { NS, parseXml } from '../../xml/index';

async function workbookFrom(sample: (typeof native.cases)[number]) {
	const workbook = createWorkbook();
	workbook.sheets[0]!.name = 'Chart data';
	workbook.theme.colors = THEME_SLOTS.map((slot) =>
		(native.scheme as Record<string, string>)[slot]!.slice(1),
	);
	createEditSession(workbook).addChart(0, {
		anchor: {
			from: { row: 0, col: 6, rowOffset: 0, colOffset: 0 },
			ext: { cx: 6000000, cy: 4000000 },
		},
		chartType: sample.type as ChartType,
		showLegend: true,
		series: [],
	});
	const zip = await JSZip.loadAsync(await saveXlsx(workbook));
	for (const [name, xml] of Object.entries(sample.parts)) zip.file(name, xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}

const chartOf = (workbook: Awaited<ReturnType<typeof workbookFrom>>) =>
	workbook.sheets[0]!.drawings[0] as ChartObject;
const axes = (xml: string) =>
	Array.from(parseXml(xml).getElementsByTagNameNS(NS.c, 'valAx')).map((node) => String(node));

describe('Change Colors against native Excel', () => {
	for (const sample of native.cases.filter((sample) =>
		['automatic', 'manual'].includes(sample.state),
	)) {
		it(`recolors ${sample.type}/${sample.state}, preserving manual fills and chart detail`, async () => {
			const expected = native.cases.find(
				(item) =>
					item.type === sample.type &&
					item.state === (sample.state === 'manual' ? 'changed' : 'automaticChanged'),
			)!;
			const workbook = await workbookFrom(sample);
			const original = structuredClone(chartOf(workbook));
			const session = createEditSession(workbook);
			const patch = chartColorPalettePatch(chartOf(workbook), 12)!;
			session.updateChart(0, 0, patch);
			const view = chartView(workbook, 0, chartOf(workbook), () => []);
			const colors = view.series.map((series) => ({
				primary: series.color.toUpperCase(),
				points: series.pointColors?.map((color) => color.toUpperCase()) ?? [],
			}));
			if (sample.type === 'pie' || sample.type === 'doughnut')
				expect(colors[0]!.points).toEqual(expected.colors[0]!.points);
			else
				expect(colors.map((color) => color.primary)).toEqual(
					expected.colors.map((color) => color.primary),
				);
			if (sample.state === 'manual' && ['column', 'bar'].includes(sample.type))
				expect(view.series[0]!.pointColors?.[0]).toBe('#00FF00');
			const saved = await saveXlsx(workbook);
			const reloaded = await loadXlsx(saved);
			expect(chartOf(reloaded).colorPalette).toBe(12);
			expect(
				chartOf(reloaded).series.map((series) => ({
					name: series.name,
					values: series.values,
					categories: series.categories,
				})),
			).toEqual(
				original.series.map((series) => ({
					name: series.name,
					values: series.values,
					categories: series.categories,
				})),
			);
			const afterView = chartView(reloaded, 0, chartOf(reloaded), () => []);
			expect(afterView.series).toEqual(view.series);
			const chartXml = await (
				await JSZip.loadAsync(saved)
			)
				.file('xl/charts/chart1.xml')!
				.async('string');
			expect(axes(chartXml)).toEqual(axes(sample.parts['xl/charts/chart1.xml']!));
			expect(chartXml).toContain('uniqueId');
			session.undo();
			expect(chartOf(workbook)).toEqual(original);
			session.redo();
			expect(chartOf(workbook).colorPalette).toBe(12);
		});
	}
	it('leaves source chart parts byte-identical before a palette edit', async () => {
		const sample = native.cases[0]!;
		const workbook = await workbookFrom(sample);
		const zip = await JSZip.loadAsync(await saveXlsx(workbook));
		for (const [name, xml] of Object.entries(sample.parts))
			expect(await zip.file(name)?.async('string')).toBe(xml);
	});
	it('lets a legacy series color edit replace the parsed DrawingML choice', async () => {
		const workbook = await workbookFrom(
			native.cases.find((item) => item.type === 'column' && item.state === 'automatic')!,
		);
		const series = structuredClone(chartOf(workbook).series);
		series[0]!.color = { rgb: '123456' };
		const session = createEditSession(workbook);
		session.updateChart(0, 0, { series });
		expect(chartOf(workbook).series[0]!.drawingColor).toBeUndefined();
		session.updateChart(0, 0, chartColorPalettePatch(chartOf(workbook), 12)!);
		const saved = await loadXlsx(await saveXlsx(workbook));
		expect(chartView(saved, 0, chartOf(saved), () => []).series[0]!.color).toBe('#123456');
	});
	it('preserves explicit point colors when a type change regenerates the chart', async () => {
		const workbook = await workbookFrom(
			native.cases.find((item) => item.type === 'pie' && item.state === 'manual')!,
		);
		const points = structuredClone(chartOf(workbook).series[0]!.pointColors);
		createEditSession(workbook).updateChart(0, 0, { chartType: 'doughnut' });
		const saved = await loadXlsx(await saveXlsx(workbook));
		expect(chartOf(saved).series[0]!.pointColors).toEqual(points);
	});
	it('copies a shared color style privately when only one chart changes', async () => {
		const sample = native.cases.find(
			(item) => item.type === 'column' && item.state === 'automatic',
		)!;
		const workbook = await workbookFrom(sample);
		const copy = structuredClone(chartOf(workbook));
		const index = createEditSession(workbook).addChart(0, copy);
		const zip = await JSZip.loadAsync(await saveXlsx(workbook));
		zip.file('xl/charts/_rels/chart2.xml.rels', sample.parts['xl/charts/_rels/chart1.xml.rels']!);
		const shared = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		createEditSession(shared).updateChart(0, 0, chartColorPalettePatch(chartOf(shared), 12)!);
		const savedBytes = await saveXlsx(shared);
		const saved = await loadXlsx(savedBytes);
		expect(chartOf(saved).colorPalette).toBe(12);
		expect((saved.sheets[0]!.drawings[index] as ChartObject).colorPalette).toBe(10);
		const after = await JSZip.loadAsync(savedBytes);
		expect(await after.file('xl/charts/chart2.xml')!.async('string')).toBe(
			await zip.file('xl/charts/chart2.xml')!.async('string'),
		);
	});
});
