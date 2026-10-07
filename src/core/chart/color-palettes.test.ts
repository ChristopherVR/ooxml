import { describe, expect, it } from 'vitest';
import native from './excel-chart-colors.json';
import custom from './excel-chart-colors-custom.json';
import { CHART_COLOR_PALETTES, chartPaletteColors, findChartColorPalette } from './color-palettes';

describe('shared chart color palettes against native Excel', () => {
	it('contains the complete Change Colors catalog', () => {
		expect(CHART_COLOR_PALETTES.map((p) => p.id)).toEqual(
			Array.from({ length: 17 }, (_, i) => i + 10),
		);
		expect(findChartColorPalette(9)).toBeUndefined();
		expect(chartPaletteColors(CHART_COLOR_PALETTES[0]!, 0, native.scheme)).toEqual([]);
	});
	for (const [theme, fixture] of Object.entries({ office: native, custom })) {
		for (const sample of fixture.cases) {
			const exact = theme === 'office' && sample.seriesCount <= 10;
			const assertion = exact ? 'matches exactly' : 'stays within one RGB channel step of';
			it(`${assertion} ${theme} palette ${sample.palette} with ${sample.seriesCount} series`, () => {
				const palette = findChartColorPalette(sample.palette)!;
				const actual = chartPaletteColors(palette, sample.seriesCount, fixture.scheme);
				if (exact) {
					expect(actual).toEqual(sample.colors);
					return;
				}
				// The extended native probes expose unresolved Office rounding at half-channel
				// boundaries. Keep the strict baseline above and measure this remaining gap.
				expect(actual).toHaveLength(sample.colors.length);
				for (let i = 0; i < actual.length; i++) {
					for (const channel of [1, 3, 5]) {
						const rendered = Number.parseInt(actual[i]!.slice(channel, channel + 2), 16);
						const nativeChannel = Number.parseInt(
							sample.colors[i]!.slice(channel, channel + 2),
							16,
						);
						expect(
							Math.abs(rendered - nativeChannel),
							`${actual[i]} versus ${sample.colors[i]}`,
						).toBeLessThanOrEqual(1);
					}
				}
			});
		}
	}
});
