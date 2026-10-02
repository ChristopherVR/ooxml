// Conditional formatting presets of Excel's Home > Conditional Formatting gallery: highlight
// formats, data bar colours, colour scales and icon sets, built as core `ConditionalRule`s.
import type { CfvoThreshold, ConditionalRule, DifferentialStyle } from '@christophervr/xlsx-core';

const argb = (hex: string) => ({ rgb: `FF${hex}` });

/** The "Format cells that..." choices of the quick rule dialogs. */
export const HIGHLIGHT_STYLES: ReadonlyArray<
	readonly [id: string, label: string, style: DifferentialStyle]
> = [
	[
		'lightRed',
		'Light Red Fill with Dark Red Text',
		{
			font: { color: argb('9C0006') },
			fill: { type: 'pattern', pattern: 'solid', bgColor: argb('FFC7CE') },
		},
	],
	[
		'yellow',
		'Yellow Fill with Dark Yellow Text',
		{
			font: { color: argb('9C5700') },
			fill: { type: 'pattern', pattern: 'solid', bgColor: argb('FFEB9C') },
		},
	],
	[
		'green',
		'Green Fill with Dark Green Text',
		{
			font: { color: argb('006100') },
			fill: { type: 'pattern', pattern: 'solid', bgColor: argb('C6EFCE') },
		},
	],
	[
		'lightRedFill',
		'Light Red Fill',
		{ fill: { type: 'pattern', pattern: 'solid', bgColor: argb('FFC7CE') } },
	],
	['redText', 'Red Text', { font: { color: argb('9C0006') } }],
	[
		'redBorder',
		'Red Border',
		{
			border: {
				left: { style: 'thin', color: argb('9C0006') },
				right: { style: 'thin', color: argb('9C0006') },
				top: { style: 'thin', color: argb('9C0006') },
				bottom: { style: 'thin', color: argb('9C0006') },
			},
		},
	],
];

export const DATA_BAR_COLORS: ReadonlyArray<readonly [id: string, label: string, hex: string]> = [
	['blue', 'Blue Data Bar', '638EC6'],
	['green', 'Green Data Bar', '63BE7B'],
	['red', 'Red Data Bar', 'F8696B'],
	['orange', 'Orange Data Bar', 'FFB628'],
	['lightBlue', 'Light Blue Data Bar', '008AEF'],
	['purple', 'Purple Data Bar', 'D6007B'],
];

export const COLOR_SCALES: ReadonlyArray<readonly [id: string, label: string, colors: string[]]> = [
	['gyr', 'Green - Yellow - Red Color Scale', ['63BE7B', 'FFEB84', 'F8696B']],
	['ryg', 'Red - Yellow - Green Color Scale', ['F8696B', 'FFEB84', '63BE7B']],
	['gwr', 'Green - White - Red Color Scale', ['63BE7B', 'FCFCFF', 'F8696B']],
	['rwg', 'Red - White - Green Color Scale', ['F8696B', 'FCFCFF', '63BE7B']],
	['bwr', 'Blue - White - Red Color Scale', ['5A8AC6', 'FCFCFF', 'F8696B']],
	['rwb', 'Red - White - Blue Color Scale', ['F8696B', 'FCFCFF', '5A8AC6']],
	['wr', 'White - Red Color Scale', ['FCFCFF', 'F8696B']],
	['rw', 'Red - White Color Scale', ['F8696B', 'FCFCFF']],
	['gw', 'Green - White Color Scale', ['63BE7B', 'FCFCFF']],
	['wg', 'White - Green Color Scale', ['FCFCFF', '63BE7B']],
	['gy', 'Green - Yellow Color Scale', ['63BE7B', 'FFEF9C']],
	['yg', 'Yellow - Green Color Scale', ['FFEF9C', '63BE7B']],
];

export const ICON_SETS: ReadonlyArray<readonly [id: string, label: string]> = [
	['3Arrows', '3 Arrows (Colored)'],
	['3ArrowsGray', '3 Arrows (Gray)'],
	['3Triangles', '3 Triangles'],
	['4Arrows', '4 Arrows (Colored)'],
	['5Arrows', '5 Arrows (Colored)'],
	['3TrafficLights1', '3 Traffic Lights (Unrimmed)'],
	['3TrafficLights2', '3 Traffic Lights (Rimmed)'],
	['3Signs', '3 Signs'],
	['4TrafficLights', '4 Traffic Lights'],
	['3Symbols', '3 Symbols (Circled)'],
	['3Symbols2', '3 Symbols (Uncircled)'],
	['3Flags', '3 Flags'],
	['3Stars', '3 Stars'],
	['4Rating', '4 Ratings'],
	['5Rating', '5 Ratings'],
	['5Quarters', '5 Quarters'],
];

export const iconCount = (set: string): number => Number(/^\d/.exec(set)?.[0] ?? 3);

/** Evenly spaced percent thresholds (`0, 33, 67` for three icons). */
export function iconThresholds(count: number): CfvoThreshold[] {
	return Array.from({ length: count }, (_v, i) => ({
		type: 'percent',
		value: String(Math.round((i * 100) / count)),
	}));
}

export function dataBarRule(hex: string): ConditionalRule {
	return {
		type: 'dataBar',
		min: { type: 'min' },
		max: { type: 'max' },
		color: argb(hex),
		priority: 1,
	};
}

export function colorScaleRule(colors: string[]): ConditionalRule {
	const thresholds: CfvoThreshold[] =
		colors.length === 3
			? [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }]
			: [{ type: 'min' }, { type: 'max' }];
	return { type: 'colorScale', thresholds, colors: colors.map(argb), priority: 1 };
}

export function iconSetRule(set: string): ConditionalRule {
	return { type: 'iconSet', iconSet: set, thresholds: iconThresholds(iconCount(set)), priority: 1 };
}

/** A CSS preview of a differential style (for the dialogs' format pickers). */
export function dxfCss(style: DifferentialStyle): string {
	const parts: string[] = [];
	const fill = style.fill;
	if (fill?.type === 'pattern') {
		const color = fill.bgColor?.rgb ?? fill.fgColor?.rgb;
		if (color) parts.push(`background:#${color.slice(-6)}`);
	}
	if (style.font?.color?.rgb) parts.push(`color:#${style.font.color.rgb.slice(-6)}`);
	if (style.font?.bold) parts.push('font-weight:bold');
	if (style.font?.italic) parts.push('font-style:italic');
	if (style.border?.top)
		parts.push(`border:1px solid #${style.border.top.color?.rgb?.slice(-6) ?? '000000'}`);
	return parts.join(';');
}
