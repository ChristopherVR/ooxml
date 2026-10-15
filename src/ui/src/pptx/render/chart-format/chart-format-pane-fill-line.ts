/**
 * The Fill & Line tab of a chart Format pane.
 *
 * @module render/chart-format/chart-format-pane-fill-line
 */
import type { ChartElementFormatView, PptxChartFormatTarget } from 'ooxml-core/pptx';

import type { ChartFormatChoice, ChartFormatSection } from './chart-format-pane-types';
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

const NOT_YET = t('notSupported', 'Not available yet in this viewer');

/** Percent 0..100 for display. */
const pct = (value: number | undefined): number => Math.round(value ?? 0);

/** `#RRGGBB` for an `<input type=color>`; unknown colours fall back to black. */
export function hexForInput(color: string | undefined, fallback = '#000000'): string {
	if (color && /^#?[0-9a-f]{6}$/iu.test(color)) {
		return `#${color.replace(/^#/u, '').toUpperCase()}`;
	}
	return fallback;
}

const DASHES: Array<[string, string]> = [
	['solid', 'Solid'],
	['sysDot', 'Round Dot'],
	['sysDash', 'Square Dot'],
	['dash', 'Dash'],
	['dashDot', 'Dash Dot'],
	['lgDash', 'Long Dash'],
	['lgDashDot', 'Long Dash Dot'],
	['lgDashDotDot', 'Long Dash Dot Dot'],
];
const COMPOUNDS: Array<[string, string]> = [
	['sng', 'Single'],
	['dbl', 'Double'],
	['thickThin', 'Thick Thin'],
	['thinThick', 'Thin Thick'],
	['tri', 'Triple'],
];
const CAPS: Array<[string, string]> = [
	['flat', 'Flat'],
	['sq', 'Square'],
	['rnd', 'Round'],
];
const JOINS: Array<[string, string]> = [
	['round', 'Round'],
	['bevel', 'Bevel'],
	['miter', 'Miter'],
];
export const ARROWS: Array<[string, string]> = [
	['none', 'None'],
	['arrow', 'Open Arrow'],
	['triangle', 'Arrow'],
	['stealth', 'Stealth Arrow'],
	['diamond', 'Diamond Arrow'],
	['oval', 'Oval Arrow'],
];

const choices = (prefix: string, list: Array<[string, string]>): ChartFormatChoice[] =>
	list.map(([value, label]) => choice(value, `${prefix}.${value}`, label));

/** The Fill section, or `undefined` for a target that is a line. */
function fillSection(view: ChartElementFormatView): ChartFormatSection | undefined {
	if (!view.support.fill) {
		return undefined;
	}
	const kind = view.fill.kind;
	const section: ChartFormatSection = {
		id: 'fill',
		title: t('section.fill', 'Fill'),
		fields: [
			{
				id: 'fill.kind',
				kind: 'radio',
				label: t('section.fill', 'Fill'),
				value: kind,
				choices: [
					choice('none', 'fill.none', 'No fill'),
					choice('solid', 'fill.solid', 'Solid fill'),
					choice('gradient', 'fill.gradient', 'Gradient fill', {
						disabled: !view.support.gradient,
						...(view.support.gradient ? {} : { hint: NOT_YET }),
					}),
					choice('picture', 'fill.picture', 'Picture or texture fill', {
						disabled: true,
						hint: NOT_YET,
					}),
					choice('pattern', 'fill.pattern', 'Pattern fill', { disabled: true, hint: NOT_YET }),
					choice('automatic', 'fill.automatic', 'Automatic'),
				],
			},
		],
	};
	if (kind === 'solid') {
		section.fields.push(
			{
				id: 'fill.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(view.fill.color),
			},
			{
				id: 'fill.transparency',
				kind: 'number',
				label: t('transparency', 'Transparency'),
				value: pct(view.fill.transparency),
				min: 0,
				max: 100,
				step: 1,
				unit: '%',
			},
		);
	} else if (kind === 'gradient' && view.fill.gradient) {
		const stops = view.fill.gradient.stops;
		section.fields.push(
			{
				id: 'fill.gradientStart',
				kind: 'color',
				label: t('gradientStart', 'Start color'),
				value: hexForInput(stops[0]?.color, '#FFFFFF'),
			},
			{
				id: 'fill.gradientEnd',
				kind: 'color',
				label: t('gradientEnd', 'End color'),
				value: hexForInput(stops[stops.length - 1]?.color),
			},
			{
				id: 'fill.gradientAngle',
				kind: 'number',
				label: t('angle', 'Angle'),
				value: Math.round(view.fill.gradient.angle ?? 0),
				min: 0,
				max: 359,
				step: 1,
				unit: '°',
			},
		);
	}
	return section;
}

/** The Border section (Line for a line target). */
function lineSection(
	view: ChartElementFormatView,
	target: PptxChartFormatTarget,
): ChartFormatSection {
	const isLine = !view.support.fill;
	const line = view.line;
	const section: ChartFormatSection = {
		id: 'line',
		title: isLine ? t('section.line', 'Line') : t('section.border', 'Border'),
		fields: [
			{
				id: 'line.kind',
				kind: 'radio',
				label: isLine ? t('section.line', 'Line') : t('section.border', 'Border'),
				value: line.kind,
				choices: [
					choice('none', 'line.none', 'No line'),
					choice('solid', 'line.solid', 'Solid line'),
					choice('gradient', 'line.gradient', 'Gradient line', { disabled: true, hint: NOT_YET }),
					choice('automatic', 'fill.automatic', 'Automatic'),
				],
			},
		],
	};
	if (line.kind === 'none') {
		return section;
	}
	if (line.kind === 'solid') {
		section.fields.push(
			{
				id: 'line.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(line.color),
			},
			{
				id: 'line.transparency',
				kind: 'number',
				label: t('transparency', 'Transparency'),
				value: pct(line.transparency),
				min: 0,
				max: 100,
				step: 1,
				unit: '%',
			},
		);
	}
	section.fields.push(
		{
			id: 'line.width',
			kind: 'number',
			label: t('width', 'Width'),
			value: line.width ?? '',
			min: 0,
			max: 1584,
			step: 0.25,
			unit: 'pt',
			auto: true,
		},
		{
			id: 'line.compound',
			kind: 'select',
			label: t('compound', 'Compound type'),
			value: line.compound ?? 'sng',
			choices: choices('compound', COMPOUNDS),
		},
		{
			id: 'line.dash',
			kind: 'select',
			label: t('dash', 'Dash type'),
			value: line.dash ?? 'solid',
			choices: choices('dash', DASHES),
		},
		{
			id: 'line.cap',
			kind: 'select',
			label: t('cap', 'Cap type'),
			value: line.cap ?? 'flat',
			choices: choices('cap', CAPS),
		},
		{
			id: 'line.join',
			kind: 'select',
			label: t('join', 'Join type'),
			value: line.join ?? 'round',
			choices: choices('join', JOINS),
		},
	);
	// Arrows apply to open lines: axes, gridlines and line series.
	if (isLine || target.kind === 'axis') {
		section.fields.push(
			{
				id: 'line.beginArrow',
				kind: 'select',
				label: t('beginArrow', 'Begin Arrow type'),
				value: line.beginArrow ?? 'none',
				choices: choices('arrow', ARROWS),
			},
			{
				id: 'line.endArrow',
				kind: 'select',
				label: t('endArrow', 'End Arrow type'),
				value: line.endArrow ?? 'none',
				choices: choices('arrow', ARROWS),
			},
		);
	}
	return section;
}

/** The Fill & Line tab's sections. */
export function fillLineSections(
	view: ChartElementFormatView,
	target: PptxChartFormatTarget,
): ChartFormatSection[] {
	const fill = fillSection(view);
	return [...(fill ? [fill] : []), lineSection(view, target)];
}
