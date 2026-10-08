import type { CommandSpec } from './ribbon-parts';
import { OFFICE_COLOR_SWATCHES } from '../form/color-swatches';

/** The common Office solid font colors, with an explicit typed action per entry. */
export function fontColorOptions(): CommandSpec[] {
	return OFFICE_COLOR_SWATCHES.map(({ hex, label }) => ({
		id: `font-color-${label.toLowerCase()}`,
		label,
		action: { type: 'font-color', value: hex },
		checked: false,
	}));
}

const colors = [
	['Black', '#000000'],
	['White', '#ffffff'],
	['Blue', '#4472c4'],
	['Orange', '#ed7d31'],
	['Green', '#70ad47'],
	['Red', '#ff0000'],
] as const;

/** Supported solid paints use the shared Office menu and source-backed format actions. */
export function paintOptions(target: 'fill' | 'line'): CommandSpec[] {
	const items: CommandSpec[] = colors.map(([label, color]) => ({
		id: `${target}-${label.toLowerCase()}`,
		label,
		action: {
			type: 'shape-format',
			patch: target === 'fill' ? { fillColor: color } : { lineColor: color },
		},
		checked: false,
	}));
	if (target === 'fill')
		items.push({
			id: 'fill-none',
			label: 'No Fill',
			action: { type: 'shape-format', patch: { fillColor: 'none' } },
			checked: false,
		});
	else
		items.push({
			id: 'line-weight',
			label: 'Weight',
			items: [0.5, 1, 1.5, 2, 3, 4, 6].map((value) => ({
				id: `line-weight-${value}`,
				label: `${value} pt`,
				action: { type: 'shape-format', patch: { lineWeight: value } },
				checked: false,
			})),
		});
	return items;
}
