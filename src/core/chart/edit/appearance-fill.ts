// Preserve authored theme tokens when a resolved host appearance changes only geometry/opacity.
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartFillAppearance, ChartLineAppearance } from './appearance-types';
import { chartFillFromAppearance, chartLineFromAppearance } from './format-values';
import { withDrawingColorTransparency } from './paint';
import type { ChartLineEdit } from './line-options';

export const sameAppearance = (a: unknown, b: unknown): boolean =>
	JSON.stringify(a) === JSON.stringify(b);

export function fillAppearancePatch(
	current: DrawingFill | undefined,
	before: ChartFillAppearance | undefined,
	after: ChartFillAppearance,
): DrawingFill | undefined {
	const converted = chartFillFromAppearance(after);
	if (!before || before.kind !== after.kind || current?.kind !== converted?.kind) return converted;
	if (current?.kind === 'solid' && converted?.kind === 'solid' && before.color === after.color)
		return {
			...current,
			color:
				(before.transparency ?? 0) === (after.transparency ?? 0)
					? current.color
					: withDrawingColorTransparency(current.color, after.transparency ?? 0),
		};
	if (current?.kind === 'gradient' && converted?.kind === 'gradient') {
		const old = before.gradient,
			next = after.gradient;
		if (
			old &&
			next &&
			old.stops.length === next.stops.length &&
			current.stops.length === next.stops.length
		)
			converted.stops = converted.stops.map((stop, index) => ({
				...current.stops[index],
				...stop,
				color:
					old.stops[index]!.color === next.stops[index]!.color
						? (old.stops[index]!.opacity ?? 1) === (next.stops[index]!.opacity ?? 1)
							? current.stops[index]!.color
							: withDrawingColorTransparency(
									current.stops[index]!.color,
									100 - (next.stops[index]!.opacity ?? 1) * 100,
								)
						: stop.color,
			}));
		return { ...current, ...converted };
	}
	if (current?.kind === 'pattern' && converted?.kind === 'pattern')
		return {
			...current,
			...converted,
			...(before.pattern?.foreground === after.pattern?.foreground && current.foreground
				? { foreground: current.foreground }
				: {}),
			...(before.pattern?.background === after.pattern?.background && current.background
				? { background: current.background }
				: {}),
		};
	return converted;
}

/** Partial outline values; null removes one property while preserving the rest of its source. */
export function lineAppearancePatch(
	current: DrawingLine | undefined,
	before: ChartLineAppearance | undefined,
	after: ChartLineAppearance,
): ChartLineEdit {
	const converted = chartLineFromAppearance(after) ?? {};
	const edit: ChartLineEdit = {};
	const fields = {
		width: 'widthEmu',
		dash: 'dash',
		cap: 'cap',
		compound: 'compound',
		join: 'join',
		beginArrow: 'headEnd',
		endArrow: 'tailEnd',
	} as const;
	for (const [source, destination] of Object.entries(fields) as [
		keyof typeof fields,
		(typeof fields)[keyof typeof fields],
	][])
		if (!sameAppearance(before?.[source], after[source]))
			Object.assign(edit, { [destination]: converted[destination] ?? null });
	const paint = (line: ChartLineAppearance | undefined): ChartFillAppearance | undefined =>
		line && {
			kind: line.kind,
			...(line.color === undefined ? {} : { color: line.color }),
			...(line.transparency === undefined ? {} : { transparency: line.transparency }),
			...(line.gradient === undefined ? {} : { gradient: line.gradient }),
		};
	if (!sameAppearance(paint(before), paint(after)))
		edit.fill = fillAppearancePatch(current?.fill, paint(before), paint(after)!) ?? null;
	return edit;
}
