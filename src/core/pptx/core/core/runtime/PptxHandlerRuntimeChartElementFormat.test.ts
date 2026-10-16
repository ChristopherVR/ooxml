import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	readChartElementFormat,
	resetChartElementFormat,
	withChartElementFormat,
} from '../../builders/sdk/chart-element-format-operations';
import {
	withChartAxisOptions,
	withChartLegendOptions,
	withChartSeriesOptions,
} from '../../builders/sdk/chart-format-options-operations';
import {
	BAR_CHART_XML,
	buildDeck,
	loadChart,
	saveWith,
	section,
} from './chart-element-format-deck.test-util';

/**
 * Load -> chart Format pane edits -> save -> reload. Every edit the panes make
 * (fill with transparency, line width/dash/compound/cap/join/arrows, shadow,
 * glow, soft edges, Reset to Match Style, Series/Axis/Legend options) must
 * land in the right element's XML and come back on reload.
 */

describe('chart Format pane edits save and reload', () => {
	it('writes legend fill, border and shadow into c:legend/c:spPr in schema order', async () => {
		const { xml } = await saveWith((data) =>
			withChartElementFormat(
				data,
				{ kind: 'legend' },
				{
					fill: { kind: 'solid', color: '#FFF2CC', transparency: 25 },
					line: { kind: 'solid', color: '#BF9000', width: 1.5, dash: 'dash' },
					effects: {
						shadow: { color: '#000000', transparency: 60, blur: 4, distance: 3, angle: 45 },
					},
				},
			),
		);
		const legend = section(xml, 'c:legend');
		expect(legend).toMatch(
			/<c:overlay val="0"\/>\s*<c:spPr><a:solidFill><a:srgbClr val="FFF2CC"><a:alpha val="75000"\/><\/a:srgbClr><\/a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="BF9000"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"/u,
		);
	});

	it('writes the chart area, series, axis line, gridlines and title border', async () => {
		const { xml, reloaded } = await saveWith((data) => {
			let next = withChartElementFormat(
				data,
				{ kind: 'chartArea' },
				{ fill: { kind: 'solid', color: '#DDEBF7', transparency: 10 } },
			);
			next = withChartElementFormat(
				next,
				{ kind: 'series', seriesIndex: 0 },
				{
					fill: { kind: 'solid', color: '#C00000' },
					effects: { glow: { color: '#FFC000', radius: 5, transparency: 40 } },
				},
			);
			next = withChartElementFormat(
				next,
				{ kind: 'axis', axisIndex: 0 },
				{ line: { kind: 'solid', color: '#7F7F7F', width: 2, endArrow: 'triangle' } },
			);
			next = withChartElementFormat(
				next,
				{ kind: 'majorGridlines', axisIndex: 1 },
				{ line: { kind: 'solid', color: '#D9D9D9', compound: 'dbl', cap: 'rnd', join: 'round' } },
			);
			return withChartElementFormat(
				next,
				{ kind: 'title' },
				{ line: { kind: 'solid', color: '#000000', width: 0.75 } },
			);
		});
		// Chart area c:spPr sits after c:chart and before c:txPr.
		expect(xml).toMatch(
			/<\/c:chart>\s*<c:spPr><a:solidFill><a:srgbClr val="DDEBF7"><a:alpha val="90000"\/>/u,
		);
		const ser = section(xml, 'c:ser');
		expect(ser).toMatch(/<\/c:tx>\s*<c:spPr><a:solidFill><a:srgbClr val="C00000"\/>/u);
		expect(ser).toMatch(/<a:effectLst><a:glow rad="63500"><a:srgbClr val="FFC000">/u);
		const catAx = section(xml, 'c:catAx');
		expect(catAx).toMatch(
			/<c:tickLblPos val="nextTo"\/>\s*<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="7F7F7F"\/><\/a:solidFill><a:tailEnd type="triangle"\/><\/a:ln><\/c:spPr>\s*<c:crossAx/u,
		);
		expect(section(xml, 'c:valAx')).toMatch(
			/<c:majorGridlines><c:spPr><a:ln cmpd="dbl" cap="rnd"><a:solidFill><a:srgbClr val="D9D9D9"\/><\/a:solidFill><a:round\/>/u,
		);
		expect(section(xml, 'c:title')).toMatch(/<c:spPr>.*<a:ln w="9525">/u);

		const chartData = reloaded.chart.chartData!;
		expect(chartData.style?.chartAreaFill?.toUpperCase()).toContain('DDEBF7');
		expect(chartData.series[0].color?.toUpperCase()).toContain('C00000');
		expect(chartData.axes?.[0].spPr?.strokeColor?.toUpperCase()).toContain('7F7F7F');
		expect(chartData.axes?.[1].majorGridlinesSpPr?.strokeColor?.toUpperCase()).toContain('D9D9D9');
	});

	it('Reset to Match Style drops the plot area formatting, effects included', async () => {
		const { xml, reloaded } = await saveWith((data) =>
			resetChartElementFormat(data, { kind: 'plotArea' }),
		);
		const plotArea = section(xml, 'c:plotArea');
		expect(plotArea).not.toContain('F2F2F2');
		expect(plotArea).not.toContain('outerShdw');
		expect(reloaded.chart.chartData!.style?.plotAreaFill).toBeUndefined();
	});

	it('persists Series, Axis and Legend options', async () => {
		const { xml, reloaded } = await saveWith((data) => {
			let next = withChartSeriesOptions(data, { gapWidth: 60, overlap: 20 });
			next = withChartAxisOptions(next, 1, {
				min: 0,
				max: 40,
				majorUnit: 10,
				numberFormat: '0.0',
				majorTickMark: 'cross',
				tickLabelPosition: 'low',
			});
			return withChartLegendOptions(next, { position: 'b', overlay: true });
		});
		expect(xml).toContain('<c:gapWidth val="60"/>');
		expect(xml).toContain('<c:overlap val="20"/>');
		expect(section(xml, 'c:legend')).toMatch(/<c:legendPos val="b"\/>\s*<c:overlay val="1"\/>/u);
		const data = reloaded.chart.chartData!;
		const valAx = data.axes?.find((a) => a.axisType === 'valAx');
		expect(valAx).toMatchObject({ min: 0, max: 40, majorUnit: 10, majorTickMark: 'cross' });
		expect(valAx?.tickLblPos).toBe('low');
		expect(valAx?.numFmt?.formatCode).toBe('0.0');
		expect(data.style?.legendPosition).toBe('b');
		expect(data.style?.legendOverlay).toBe(true);
		expect(data.barGapWidth).toBe(60);
	});

	it('leaves an untouched chart part byte-for-byte', async () => {
		const { handler, data } = await loadChart(await buildDeck());
		data.slides[0].isDirty = true;
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toBe(BAR_CHART_XML);
	});

	it('reads the formats back on reopen so the pane shows them', async () => {
		const { reloaded } = await saveWith((data) => {
			let next = withChartElementFormat(
				data,
				{ kind: 'legend' },
				{
					fill: { kind: 'solid', color: '#FFF2CC', transparency: 25 },
					line: { kind: 'solid', color: '#BF9000', width: 1.5, dash: 'dash', compound: 'dbl' },
					effects: {
						shadow: { color: '#000000', transparency: 60, blur: 4, distance: 3, angle: 45 },
						softEdge: 2,
					},
				},
			);
			next = withChartElementFormat(
				next,
				{ kind: 'series', seriesIndex: 0 },
				{ fill: { transparency: 30 }, effects: { glow: { color: '#FFC000', radius: 5 } } },
			);
			return withChartElementFormat(
				next,
				{ kind: 'axis', axisIndex: 0 },
				{ line: { kind: 'solid', color: '#7F7F7F', endArrow: 'triangle', cap: 'rnd' } },
			);
		});
		const data = reloaded.chart.chartData!;
		const legend = readChartElementFormat(data, { kind: 'legend' });
		expect(legend.fill).toMatchObject({ kind: 'solid', transparency: 25 });
		expect(legend.fill.color?.toUpperCase()).toBe('#FFF2CC');
		expect(legend.line).toMatchObject({ kind: 'solid', width: 1.5, dash: 'dash', compound: 'dbl' });
		expect(legend.effects.shadow).toMatchObject({
			transparency: 60,
			blur: 4,
			distance: 3,
			angle: 45,
		});
		expect(legend.effects.softEdge).toBe(2);
		const series = readChartElementFormat(data, { kind: 'series', seriesIndex: 0 });
		expect(series.fill).toMatchObject({ kind: 'solid', transparency: 30 });
		expect(series.effects.glow).toMatchObject({ radius: 5 });
		expect(readChartElementFormat(data, { kind: 'axis', axisIndex: 0 }).line).toMatchObject({
			kind: 'solid',
			endArrow: 'triangle',
			cap: 'rnd',
		});
		// The plot area's authored fill and shadow are read too.
		expect(data.elementFormats?.plotArea?.effects?.shadow).toBeDefined();
	});

	it('keeps authored formats untouched when only another part of the chart is edited', async () => {
		const { xml } = await saveWith((data) => ({ ...data, title: 'Renamed' }));
		// The series keeps its theme colour and the plot area its shadow, as authored.
		expect(section(xml, 'c:ser')).toContain('<a:schemeClr val="accent1"/>');
		expect(section(xml, 'c:plotArea')).toContain('<a:outerShdw blurRad="1"/>');
	});
});
