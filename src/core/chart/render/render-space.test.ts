import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { ChartSpace } from '../model';
import { parseChartSpace } from '../parse-space';
import { chartSpaceKind, chartThemePalette, renderChartSpaceSvg } from './render-space';

async function excelChartSpace(): Promise<ChartSpace> {
	const zip = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '../../xlsx/__fixtures__/excel-features.xlsx')),
	);
	return parseChartSpace(await zip.file('xl/charts/chart1.xml')!.async('string')).chartSpace;
}

/** The shape drawn inside each point group of a series (`<g data-chart-series>`). */
const marks = (svg: string, series: number) =>
	svg.match(new RegExp(`<g data-chart-series="${series}"[^>]*><[a-z]+ [^>]*>`, 'g')) ?? [];

describe('renderChartSpaceSvg', () => {
	it('draws the series of an Excel chart part from its cached values', async () => {
		const space = await excelChartSpace();
		const svg = renderChartSpaceSvg(space, { width: 480, height: 288 });
		expect(svg).toMatch(/^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" width="480" height="288"/);
		expect(svg).toContain('<title>Sales by region</title>');
		// Two series (Sales, Cost) of four regions: one bar per point, in the theme accents.
		expect(marks(svg, 0)).toHaveLength(4);
		expect(marks(svg, 1)).toHaveLength(4);
		expect(marks(svg, 0).every((mark) => mark.includes('fill="#4472C4"'))).toBe(true);
		expect(marks(svg, 1).every((mark) => mark.includes('fill="#ED7D31"'))).toBe(true);
		for (const label of ['Sales', 'Cost', 'North', 'West']) expect(svg).toContain(`>${label}<`);
		// Taller values draw taller bars: West (175) above North (100).
		const heights = marks(svg, 0).map((mark) => Number(/height="([\d.]+)"/.exec(mark)?.[1]));
		expect(heights[3]!).toBeGreaterThan(heights[0]!);
	});

	it('takes scheme colours and fonts from the document theme', async () => {
		const space = await excelChartSpace();
		const svg = renderChartSpaceSvg(space, {
			width: 480,
			height: 288,
			theme: { colors: { accent1: { kind: 'srgb', value: '123456', transforms: [] } } },
			fonts: { minor: 'Georgia' },
		});
		expect(marks(svg, 0).every((mark) => mark.includes('fill="#123456"'))).toBe(true);
		expect(svg).toContain('Georgia');
		expect(marks(svg, 1).every((mark) => mark.includes('fill="#ED7D31"'))).toBe(true);
	});

	it('fills theme slots the scheme lacks with the Office defaults', () => {
		const palette = chartThemePalette({
			colors: { accent2: { kind: 'srgb', value: 'abcdef', transforms: [] } },
		});
		expect(palette.colors).toHaveLength(12);
		expect(palette.colors[4]).toBe('4472C4');
		expect(palette.colors[5]).toBe('ABCDEF');
		expect(palette).toMatchObject({ majorFont: 'Calibri Light', minorFont: 'Calibri' });
	});

	it('reports which chart families the painter draws', async () => {
		const space = await excelChartSpace();
		expect(chartSpaceKind(space)).toEqual({ type: 'column', drawn: true });
		const bubble = structuredClone(space);
		bubble.plotArea.groups[0]!.kind = 'bubble';
		expect(chartSpaceKind(bubble)).toEqual({ type: 'bubble', drawn: false });
		expect(chartSpaceKind({ plotArea: { groups: [], axes: [] } })).toEqual({ drawn: false });
	});
});
