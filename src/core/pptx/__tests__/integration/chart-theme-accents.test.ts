/**
 * A chart with no colour-style part paints its `c:style` palette over the deck
 * theme's accents (PowerPoint does; the renderer reads them from
 * `chartData.themeAccentColors`), so load must record the theme's accent1 to
 * accent6 on every chart, and save must not write them anywhere.
 *
 * The fixture is a snapshot of `e2e/pptx/fixtures/chart-style-palette.pptx`
 * (`generate-chart-style-palette-fixture.ts`): one style-2 bar chart, no colour
 * part, theme accent1 `#2E86AB`.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';

const fixture = fileURLToPath(new URL('../fixtures/e2e/chart-style-palette.pptx', import.meta.url));

async function load(): Promise<{
	handler: PptxHandler;
	data: Awaited<ReturnType<PptxHandler['load']>>;
}> {
	const bytes = readFileSync(fixture);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
	);
	return { handler, data };
}

describe('chart theme accents', () => {
	it('records the theme accent1 to accent6 on a chart with no colour-style part', async () => {
		const { data } = await load();
		const chart = data.slides.flatMap((s) => s.elements).find((e) => e.type === 'chart');
		expect(chart?.type === 'chart' ? chart.chartData?.colorPalette : 'missing').toBeUndefined();
		const accents = chart?.type === 'chart' ? chart.chartData?.themeAccentColors : undefined;
		expect(accents).toHaveLength(6);
		expect(accents?.[0]).toBe('#2E86AB');
		expect(accents?.every((a) => /^#[0-9A-F]{6}$/u.test(a))).toBe(true);
	});

	it('writes no accent colours into the saved chart part', async () => {
		const { handler, data } = await load();
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const chartXml = await zip.file('ppt/charts/chart1.xml')?.async('string');
		expect(chartXml).toBeDefined();
		expect(chartXml).not.toMatch(/2E86AB/iu);
	});
});
