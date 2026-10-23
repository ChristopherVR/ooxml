/**
 * The Text Options half of a chart Format pane (Text Fill & Outline, Text
 * Effects and Textbox), for the chart area, title, legend, axis labels and
 * data labels. Field ids start with `text.`. An extension: only a source that
 * reads text formatting offers it.
 *
 * @module chart/format-pane-text
 */
import { choice, hexForInput, pct, t } from './format-pane-common';
import { glowSection, shadowSection } from './format-pane-effects';
import { gradientFields } from './format-pane-fill-kinds';
import type { ChartFormatField, ChartFormatSection, ChartFormatTab } from './format-pane-types';
import type { ChartFormatTextView } from './format-pane-view';

/** The Text Options tabs, in PowerPoint's order. */
export function chartTextTabs(): ChartFormatTab[] {
	return [
		{ id: 'textFill', label: t('tab.textFill', 'Text Fill & Outline') },
		{ id: 'textEffects', label: t('tab.textEffects', 'Text Effects') },
		{ id: 'textBox', label: t('tab.textBox', 'Textbox') },
	];
}

const transparency = (id: string, value: number | undefined): ChartFormatField => ({
	id,
	kind: 'number',
	label: t('transparency', 'Transparency'),
	value: pct(value),
	min: 0,
	max: 100,
	step: 1,
	unit: '%',
});

function textFillSections(text: ChartFormatTextView): ChartFormatSection[] {
	const fill = text.fill ?? { kind: 'automatic' as const };
	const fillSection: ChartFormatSection = {
		id: 'textFill',
		title: t('section.textFill', 'Text Fill'),
		fields: [
			{
				id: 'text.fill.kind',
				kind: 'radio',
				label: t('section.textFill', 'Text Fill'),
				value: fill.kind,
				choices: [
					choice('none', 'fill.none', 'No fill'),
					choice('solid', 'fill.solid', 'Solid fill'),
					choice('gradient', 'fill.gradient', 'Gradient fill'),
					choice('automatic', 'fill.automatic', 'Automatic'),
				],
			},
		],
	};
	if (fill.kind === 'solid') {
		fillSection.fields.push(
			{
				id: 'text.fill.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(fill.color),
			},
			transparency('text.fill.transparency', fill.transparency),
		);
	} else if (fill.kind === 'gradient' && fill.gradient) {
		fillSection.fields.push(...gradientFields('text.fill', fill.gradient));
	}
	const outline = text.outline ?? { kind: 'automatic' as const };
	const outlineSection: ChartFormatSection = {
		id: 'textOutline',
		title: t('section.textOutline', 'Text Outline'),
		fields: [
			{
				id: 'text.outline.kind',
				kind: 'radio',
				label: t('section.textOutline', 'Text Outline'),
				value: outline.kind,
				choices: [
					choice('none', 'line.none', 'No line'),
					choice('solid', 'line.solid', 'Solid line'),
					choice('automatic', 'fill.automatic', 'Automatic'),
				],
			},
		],
	};
	if (outline.kind === 'solid') {
		outlineSection.fields.push(
			{
				id: 'text.outline.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(outline.color),
			},
			transparency('text.outline.transparency', outline.transparency),
			{
				id: 'text.outline.width',
				kind: 'number',
				label: t('width', 'Width'),
				value: outline.width ?? 0.75,
				min: 0,
				max: 1584,
				step: 0.25,
				unit: 'pt',
			},
		);
	}
	return [fillSection, outlineSection];
}

function textBoxSection(text: ChartFormatTextView): ChartFormatSection {
	return {
		id: 'textBox',
		title: t('section.textBox', 'Text Box'),
		fields: [
			{
				id: 'text.anchor',
				kind: 'select',
				label: t('verticalAlignment', 'Vertical alignment'),
				value: text.anchor ?? 'ctr',
				choices: [
					choice('t', 'anchor.top', 'Top'),
					choice('ctr', 'anchor.middle', 'Middle'),
					choice('b', 'anchor.bottom', 'Bottom'),
				],
			},
			{
				id: 'text.direction',
				kind: 'select',
				label: t('textDirection', 'Text direction'),
				value: text.direction ?? 'horz',
				choices: [
					choice('horz', 'direction.horizontal', 'Horizontal'),
					choice('vert', 'direction.rotate90', 'Rotate all text 90°'),
					choice('vert270', 'direction.rotate270', 'Rotate all text 270°'),
					choice('wordArtVert', 'direction.stacked', 'Stacked'),
				],
			},
			{
				id: 'text.rotation',
				kind: 'number',
				label: t('customAngle', 'Custom angle'),
				value: text.rotation ?? '',
				min: -90,
				max: 90,
				step: 1,
				unit: '°',
				auto: true,
			},
		],
	};
}

/** The sections of Text Options tab `tab`. */
export function textOptionsSections(
	text: ChartFormatTextView,
	tab: 'textFill' | 'textEffects' | 'textBox',
): ChartFormatSection[] {
	if (tab === 'textEffects') {
		return [shadowSection('text.', text.effects?.shadow), glowSection('text.', text.effects?.glow)];
	}
	return tab === 'textBox' ? [textBoxSection(text)] : textFillSections(text);
}
