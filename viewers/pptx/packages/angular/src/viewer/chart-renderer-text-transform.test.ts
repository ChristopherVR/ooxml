/**
 * Source-wiring regression test for `ChartRendererComponent`'s text blocks.
 * No TestBed template mount here (see `chart-renderer.component.test.ts`), so
 * this asserts the template binds every `SvgText` field the shared view model
 * relies on. Sunburst ring labels are rotated through `SvgText.transform`; the
 * data-label block dropped it, so Angular painted them unrotated while React
 * and vanilla rotated them (caught by `e2e/chart-svg-parity.spec.ts`).
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

const source = readFileSync(path.join(__dirname, 'chart-renderer.component.ts'), 'utf8');

function textBlock(loopVar: string, list: string): string {
	const start = source.indexOf(`@for (${loopVar} of vm().${list}`);
	expect(start).toBeGreaterThan(-1);
	return source.slice(start, source.indexOf('</text>', start));
}

describe('chartRendererComponent text wiring', () => {
	it.each([
		['dl', 'dataLabels'],
		['lbl', 'axisLabels'],
		['lbl', 'secondaryAxisLabels'],
		['lbl', 'categoryLabels'],
	])('binds transform on %s of %s', (loopVar, list) => {
		expect(textBlock(loopVar, list)).toContain(`[attr.transform]="${loopVar}.transform`);
	});

	// The view model resolves the chart text font (`chartTextFontFamily`) onto axis,
	// category and data labels; Angular, Vue and Svelte dropped it, so their labels
	// fell back to the page font while React and vanilla drew Calibri (caught by
	// `e2e/pptx/chart-svg-parity.spec.ts`).
	it.each([
		['dl', 'dataLabels'],
		['lbl', 'axisLabels'],
		['lbl', 'secondaryAxisLabels'],
		['lbl', 'categoryLabels'],
	])('binds the font weight, style and family on %s of %s', (loopVar, list) => {
		const block = textBlock(loopVar, list);
		expect(block).toContain(`[attr.font-weight]="${loopVar}.fontWeight ?? 'normal'"`);
		expect(block).toContain(`[attr.font-style]="${loopVar}.fontStyle ?? 'normal'"`);
		expect(block).toContain(`[attr.font-family]="${loopVar}.fontFamily ?? null"`);
	});

	it('binds opacity on data labels', () => {
		expect(textBlock('dl', 'dataLabels')).toContain('[attr.opacity]="dl.opacity ?? 1"');
	});
});
