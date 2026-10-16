/**
 * Caption and colour helpers every chart Format pane section shares.
 *
 * @module render/chart-format/chart-format-pane-common
 */
import type { ChartFormatChoice } from './chart-format-pane-types';
import type { ChartFormatText } from './chart-format-targets';

export const t = (key: string, fallback: string): ChartFormatText => ({
	key: `pptx.chartFormat.${key}`,
	fallback,
});

export const choice = (
	value: string,
	key: string,
	fallback: string,
	extra?: Partial<ChartFormatChoice>,
): ChartFormatChoice => ({ value, label: t(key, fallback), ...extra });

/** Percent 0..100 for display. */
export const pct = (value: number | undefined): number => Math.round(value ?? 0);

/** `#RRGGBB` for an `<input type=color>`; unknown colours fall back to black. */
export function hexForInput(color: string | undefined, fallback = '#000000'): string {
	if (color && /^#?[0-9a-f]{6}$/iu.test(color)) {
		return `#${color.replace(/^#/u, '').toUpperCase()}`;
	}
	return fallback;
}
