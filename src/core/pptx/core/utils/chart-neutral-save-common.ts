import type { PptxChartData, PptxChartStyle } from '../types';

/** JSON of a value with sorted keys, `undefined` members and empty objects left out. */
function canonical(value: unknown): unknown {
	if (Array.isArray(value)) return value.map((item) => canonical(item) ?? null);
	if (value && typeof value === 'object') {
		const entries = Object.keys(value)
			.sort()
			.flatMap((key) => {
				const item = canonical((value as Record<string, unknown>)[key]);
				return item === undefined ? [] : [[key, item] as const];
			});
		return entries.length > 0 ? Object.fromEntries(entries) : undefined;
	}
	return value;
}

export const same = (a: unknown, b: unknown): boolean =>
	JSON.stringify(canonical(a)) === JSON.stringify(canonical(b));

type StyleKey = keyof PptxChartStyle;

/** Copies `keys` of `before.style` onto `probe.style` (the fields a family has claimed). */
export function revertStyle(
	probe: PptxChartData,
	before: PptxChartData,
	keys: readonly StyleKey[],
): void {
	const style: Record<string, unknown> = { ...probe.style };
	for (const key of keys) {
		const value = before.style?.[key];
		if (value === undefined) delete style[key];
		else style[key] = value;
	}
	probe.style = style as PptxChartStyle;
}
