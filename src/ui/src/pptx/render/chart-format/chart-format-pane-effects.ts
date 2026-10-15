/**
 * The Effects tab of a chart Format pane: Shadow, Glow, Soft Edges and the
 * 3-D Format section (reported as not available yet).
 *
 * @module render/chart-format/chart-format-pane-effects
 */
import type { ChartElementFormatView } from 'ooxml-core/pptx';

import { hexForInput, t } from './chart-format-pane-fill-line';
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

/** The Effects tab's sections. */
export function effectsSections(view: ChartElementFormatView): ChartFormatSection[] {
	const shadow = view.effects.shadow ?? undefined;
	const glow = view.effects.glow ?? undefined;
	const shadowSection: ChartFormatSection = {
		id: 'shadow',
		title: t('section.shadow', 'Shadow'),
		fields: [
			{
				id: 'shadow.enabled',
				kind: 'checkbox',
				label: t('shadow.enabled', 'Outer shadow'),
				value: Boolean(shadow),
			},
		],
	};
	if (shadow) {
		shadowSection.fields.push(
			{
				id: 'shadow.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(shadow.color),
			},
			num(
				'shadow.transparency',
				'transparency',
				'Transparency',
				shadow.transparency ?? 0,
				100,
				'%',
			),
			num('shadow.blur', 'blur', 'Blur', shadow.blur, 100, 'pt', 0.5),
			num('shadow.angle', 'angle', 'Angle', shadow.angle, 359, '°'),
			num('shadow.distance', 'distance', 'Distance', shadow.distance, 200, 'pt', 0.5),
		);
	}
	const glowSection: ChartFormatSection = {
		id: 'glow',
		title: t('section.glow', 'Glow'),
		fields: [
			{
				id: 'glow.enabled',
				kind: 'checkbox',
				label: t('glow.enabled', 'Glow'),
				value: Boolean(glow),
			},
		],
	};
	if (glow) {
		glowSection.fields.push(
			{
				id: 'glow.color',
				kind: 'color',
				label: t('color', 'Color'),
				value: hexForInput(glow.color),
			},
			num('glow.radius', 'size', 'Size', glow.radius, 150, 'pt', 0.5),
			num('glow.transparency', 'transparency', 'Transparency', glow.transparency ?? 0, 100, '%'),
		);
	}
	return [
		shadowSection,
		glowSection,
		{
			id: 'softEdges',
			title: t('section.softEdges', 'Soft Edges'),
			fields: [num('softEdge.size', 'size', 'Size', view.effects.softEdge ?? 0, 100, 'pt', 0.5)],
		},
		{
			id: 'threeD',
			title: t('section.threeD', '3-D Format'),
			fields: [],
			note: t('notSupported', 'Not available yet in this viewer'),
		},
	];
}
