/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * A chart with `<c:style val="2"/>`, no explicit series colours and no colour
 * style part takes its series fills from Office's chart-style gallery over the
 * deck theme: series 1 to 6 are accent1 to accent6, series 7 a darker accent1
 * (`getChartStylePalette` in `src/ui/src/pptx/render/chart-helpers.ts`, pinned
 * by its unit tests). This spec is the browser check that every binding paints
 * what the shared palette decided rather than a fallback of its own.
 *
 * The fixture (`fixtures/generate-chart-style-palette-fixture.ts`) sets the
 * theme's accent1 to a colour no built-in palette uses, so a hard-coded Office
 * palette cannot pass by coincidence. One category and seven series put the
 * seven bars left to right in series order.
 */
import { expect, test } from '@playwright/test';

import {
	CHART_STYLE_PALETTE_ACCENT1,
	CHART_STYLE_PALETTE_SERIES,
	CHART_STYLE_PALETTE_TITLE,
} from './fixtures/generate-chart-style-palette-fixture';
import { colorsMatch } from './support/color-match';
import { fixture, loadDeck } from './support/deck';
import { fingerprintCharts } from './support/svg-fingerprint';
import type { SvgPrimitiveShape } from './support/svg-fingerprint';

test.use({ viewport: { width: 1440, height: 900 } });

function rgbOf(hex: string): string {
	const n = Number.parseInt(hex, 16);
	return `rgb(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`;
}

/** Relative luminance (sRGB, 0 to 1) of a computed `rgb()` colour. */
function luminance(color: string): number {
	const parts = /rgba?\(([^)]+)\)/u
		.exec(color)?.[1]
		?.split(',')
		.map((part) => Number.parseFloat(part) / 255);
	if (!parts || parts.length < 3) {
		return Number.NaN;
	}
	const [r, g, b] = parts.map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
	return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
}

const isPainted = (fill: string): boolean =>
	fill.startsWith('rgb') && !/rgba\([^)]*,\s*0\)/u.test(fill);

/**
 * The seven bar marks, left to right: painted rects or paths sharing one width
 * (axes are lines, and the plot and chart areas are wider than any bar).
 */
function barsOf(shapes: SvgPrimitiveShape[]): SvgPrimitiveShape[] {
	const groups = new Map<number, SvgPrimitiveShape[]>();
	for (const shape of shapes) {
		if ((shape.tag !== 'rect' && shape.tag !== 'path') || !isPainted(shape.fill)) {
			continue;
		}
		const [, , width = 0, height = 0] = shape.geometry;
		if (width <= 0 || height <= 0) {
			continue;
		}
		const key = Math.round(width);
		groups.set(key, [...(groups.get(key) ?? []), shape]);
	}
	const bars = [...groups.values()].find((group) => group.length === CHART_STYLE_PALETTE_SERIES);
	return [...(bars ?? [])].sort((a, b) => (a.geometry[0] ?? 0) - (b.geometry[0] ?? 0));
}

test('a style-2 chart with no colour part paints the theme accents, then darker shades', async ({
	page,
}) => {
	await loadDeck(page, fixture('chart-style-palette.pptx'));

	let bars: SvgPrimitiveShape[] = [];
	await expect
		.poll(async () => {
			const chart = (await fingerprintCharts(page)).find((candidate) =>
				candidate.texts.some((text) => text.text === CHART_STYLE_PALETTE_TITLE),
			);
			bars = chart ? barsOf(chart.shapes) : [];
			return bars.length;
		})
		.toBe(CHART_STYLE_PALETTE_SERIES);

	const fills = bars.map((bar) => bar.fill);
	expect(
		colorsMatch(fills[0]!, rgbOf(CHART_STYLE_PALETTE_ACCENT1), 2),
		`series 1 is the theme's accent1 (#${CHART_STYLE_PALETTE_ACCENT1}), got ${fills[0]}`,
	).toBe(true);
	expect(new Set(fills.slice(0, 6)).size, `series 1 to 6 are six accents: ${fills.join(' ')}`).toBe(
		6,
	);
	expect(
		luminance(fills[6]!),
		`series 7 (${fills[6]}) is darker than series 1 (${fills[0]})`,
	).toBeLessThan(luminance(fills[0]!));
});
