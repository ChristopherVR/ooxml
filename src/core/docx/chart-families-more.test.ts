// Bubble, stock, treemap and sunburst charts authored by Word (`generate-word-chart-families.ps1`):
// the classic ones are `c:` parts, the hierarchies extended (`cx:`) parts with cached levels, all
// painted by the shared chart painter.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx, type DocxChart } from './index';
import { docxChartPaint } from './chart-paint';

const bytes = () =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__/chart-families.docx')));

async function load() {
	const loaded = await loadDocx(bytes());
	const charts = chartsIn(loaded.model.blocks);
	const by = (title: string): DocxChart => charts.find((c) => c.title === title)!;
	const paint = (title: string) => docxChartPaint(by(title), 576, 336, loaded.model.theme);
	return { loaded, charts, by, paint };
}

describe('Word bubble chart', () => {
	it('draws a translucent circle per point, sized by the cached bubble sizes', async () => {
		const { paint } = await load();
		const result = paint('Bubbles');
		expect(result.label).toBe('Chart: Bubbles');
		const svg = result.svg!;
		const radii = [...svg.matchAll(/<circle [^>]*\br="([\d.]+)"[^>]*fill-opacity="0.6"/g)].map(
			(m) => Number(m[1]),
		);
		expect(radii).toHaveLength(4);
		// Sizes 5, 10, 3, 8 by point.
		expect(radii[1]).toBeGreaterThan(radii[3]!);
		expect(radii[3]).toBeGreaterThan(radii[0]!);
		expect(radii[0]).toBeGreaterThan(radii[2]!);
		expect(svg).not.toContain('not drawn');
	});
});

describe('Word stock chart', () => {
	it('draws open-high-low-close candles with up and down bodies', async () => {
		const { paint } = await load();
		const result = paint('Prices');
		expect(result.label).toBe('Chart: Prices');
		const candles = result.svg!.match(/<g data-chart-series="3"[\s\S]*?<\/g>/g)!;
		expect(candles).toHaveLength(4);
		expect(candles.map((g) => /<rect [^>]*fill="(#[0-9A-F]+)"/.exec(g)![1])).toEqual([
			'#FFFFFF',
			'#FFFFFF',
			'#404040',
			'#FFFFFF',
		]);
		for (const name of ['Open', 'High', 'Low', 'Close']) expect(result.svg).toContain(`>${name}<`);
	});
});

describe('Word treemap and sunburst charts', () => {
	it('reads the cached category levels of the treemap', async () => {
		const { by } = await load();
		const chart = by('Produce treemap');
		expect(chart).toMatchObject({ layoutId: 'treemap' });
		expect(chart.chartEx?.hierarchy?.levels).toEqual([
			['Apple', 'Pear', 'Leek', 'Kale', 'Almond'],
			['Fruit', 'Fruit', 'Veg', 'Veg', 'Nuts'],
		]);
		expect(chart.issues).toEqual([]);
	});

	it('draws a cell per leaf with a branch legend and parent labels', async () => {
		const { paint } = await load();
		const result = paint('Produce treemap');
		expect(result.label).toBe('Chart: Produce treemap');
		const svg = result.svg!;
		expect(svg.match(/<g data-chart-series="0" data-chart-point="\d"/g)).toHaveLength(5);
		for (const label of ['Fruit', 'Veg', 'Nuts']) expect(svg).toContain(`>${label}<`);
		expect(svg).not.toContain('not drawn');
	});

	it('draws the sunburst with a ring per level', async () => {
		const { paint } = await load();
		const result = paint('Produce sunburst');
		expect(result.label).toBe('Chart: Produce sunburst');
		expect(result.svg!.match(/<path d="M [^"]*" fill="/g)).toHaveLength(8);
		for (const label of ['Fruit', 'Apple', 'Almond'])
			expect(result.svg).toContain(`>${label}</text>`);
	});
});
