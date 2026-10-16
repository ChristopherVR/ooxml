/**
 * The Fill & Line tab of a chart Format pane.
 *
 * @module render/chart-format/chart-format-pane-fill-line
 */
import type { ChartElementFormatView, PptxChartFormatTarget } from 'ooxml-core/pptx';

import { choice, hexForInput, pct, t } from './chart-format-pane-common';
import { gradientFields, patternFields, pictureFields } from './chart-format-pane-fill-kinds';
import type { ChartFormatChoice, ChartFormatSection } from './chart-format-pane-types';

export { choice, hexForInput, t } from './chart-format-pane-common';

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
					choice('gradient', 'fill.gradient', 'Gradient fill'),
					choice('picture', 'fill.picture', 'Picture or texture fill'),
					choice('pattern', 'fill.pattern', 'Pattern fill'),
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
		section.fields.push(...gradientFields('fill', view.fill.gradient));
	} else if (kind === 'pattern' && view.fill.pattern) {
		section.fields.push(...patternFields(view.fill.pattern));
	} else if (kind === 'picture') {
		section.fields.push(...pictureFields(view.fill.picture));
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
					choice('gradient', 'line.gradient', 'Gradient line'),
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
	} else if (line.kind === 'gradient' && line.gradient) {
		section.fields.push(...gradientFields('line', line.gradient));
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
