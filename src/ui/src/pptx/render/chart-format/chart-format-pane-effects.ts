/**
 * The Effects tab of a chart Format pane: Shadow, Glow, Soft Edges and, for
 * elements that have one, the 3-D Format section.
 *
 * @module render/chart-format/chart-format-pane-effects
 */
import type { ChartElementFormatView } from 'ooxml-core/pptx';

import { threeDSection } from './chart-format-pane-3d';
import { hexForInput, t } from './chart-format-pane-common';
import type { ChartFormatSection } from './chart-format-pane-types';

/** PowerPoint's "Outer: Offset Bottom Right" preset, what ticking Shadow applies. */
export const DEFAULT_CHART_SHADOW = {
	color: '#000000',
	transparency: 60,
	blur: 4,
	distance: 3,
	angle: 45,
} as const;

/** PowerPoint's "Glow: 5 point; Accent color 1" preset, with a neutral colour. */
export const DEFAULT_CHART_GLOW = { color: '#4472C4', transparency: 60, radius: 5 } as const;

const num = (
	id: string,
	key: string,
	fallback: string,
	value: number,
	max: number,
	unit: string,
	step = 1,
) => ({
	id,
	kind: 'number' as const,
	label: t(key, fallback),
	value: Math.round(value * 100) / 100,
	min: 0,
	max,
	step,
	unit,
});

type Effects = ChartElementFormatView['effects'];

/** The Shadow section; field ids start with `${prefix}shadow.` (`text.` for Text Effects). */
export function shadowSection(prefix: string, shadow: Effects['shadow']): ChartFormatSection {
	const section: ChartFormatSection = {
		id: `${prefix}shadow`,
		title: t('section.shadow', 'Shadow'),
		fields: [
			{
				id: `${prefix}shadow.enabled`,
				kind: 'checkbox',
				label: t('shadow.enabled', 'Outer shadow'),
				value: Boolean(shadow),
			},
		],
	};
	if (shadow) {
		section.fields.push(
			{
				id: `${prefix}shadow.color`,
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(shadow.color),
			},
			num(
				`${prefix}shadow.transparency`,
				'transparency',
				'Transparency',
				shadow.transparency ?? 0,
				100,
				'%',
			),
			num(`${prefix}shadow.blur`, 'blur', 'Blur', shadow.blur, 100, 'pt', 0.5),
			num(`${prefix}shadow.angle`, 'angle', 'Angle', shadow.angle, 359, '°'),
			num(`${prefix}shadow.distance`, 'distance', 'Distance', shadow.distance, 200, 'pt', 0.5),
		);
	}
	return section;
}

/** The Glow section; field ids start with `${prefix}glow.`. */
export function glowSection(prefix: string, glow: Effects['glow']): ChartFormatSection {
	const section: ChartFormatSection = {
		id: `${prefix}glow`,
		title: t('section.glow', 'Glow'),
		fields: [
			{
				id: `${prefix}glow.enabled`,
				kind: 'checkbox',
				label: t('glow.enabled', 'Glow'),
				value: Boolean(glow),
			},
		],
	};
	if (glow) {
		section.fields.push(
			{
				id: `${prefix}glow.color`,
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(glow.color),
			},
			num(`${prefix}glow.radius`, 'size', 'Size', glow.radius, 150, 'pt', 0.5),
			num(
				`${prefix}glow.transparency`,
				'transparency',
				'Transparency',
				glow.transparency ?? 0,
				100,
				'%',
			),
		);
	}
	return section;
}

/** The Effects tab's sections. */
export function effectsSections(view: ChartElementFormatView): ChartFormatSection[] {
	return [
		shadowSection('', view.effects.shadow),
		glowSection('', view.effects.glow),
		{
			id: 'softEdges',
			title: t('section.softEdges', 'Soft Edges'),
			fields: [num('softEdge.size', 'size', 'Size', view.effects.softEdge ?? 0, 100, 'pt', 0.5)],
		},
		...(view.support.threeD ? [threeDSection(view.threeD)] : []),
	];
}
