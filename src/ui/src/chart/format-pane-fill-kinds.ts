/**
 * The detail fields of the gradient, pattern and picture choices of a chart
 * Format pane's Fill, Line and Text Fill sections.
 *
 * @module chart/format-pane-fill-kinds
 */
import type { ChartGradientFill } from 'ooxml-core/chart';

import { choice, hexForInput, t } from './format-pane-common';
import type { ChartFormatChoice, ChartFormatField } from './format-pane-types';
import type { ChartFormatFillView } from './format-pane-view';

/** Type, stops and angle of a gradient; field ids are `${prefix}.gradient*`. */
export function gradientFields(prefix: string, gradient: ChartGradientFill): ChartFormatField[] {
	const stops = gradient.stops;
	const fields: ChartFormatField[] = [
		{
			id: `${prefix}.gradientType`,
			kind: 'select',
			label: t('gradientType', 'Type'),
			value: gradient.type,
			choices: [
				choice('linear', 'gradientType.linear', 'Linear'),
				choice('radial', 'gradientType.radial', 'Radial'),
			],
		},
		{
			id: `${prefix}.gradientStart`,
			kind: 'color',
			label: t('gradientStart', 'Start color'),
			value: hexForInput(stops[0]?.color, '#FFFFFF'),
		},
		{
			id: `${prefix}.gradientEnd`,
			kind: 'color',
			label: t('gradientEnd', 'End color'),
			value: hexForInput(stops[stops.length - 1]?.color),
		},
	];
	if (gradient.type === 'linear') {
		fields.push({
			id: `${prefix}.gradientAngle`,
			kind: 'number',
			label: t('angle', 'Angle'),
			value: Math.round(gradient.angle ?? 0),
			min: 0,
			max: 359,
			step: 1,
			unit: '°',
		});
	}
	return fields;
}

/** Preset (one of `presets`), foreground and background of a pattern fill. */
export function patternFields(
	pattern: NonNullable<ChartFormatFillView['pattern']>,
	presets: readonly ChartFormatChoice[],
): ChartFormatField[] {
	return [
		{
			id: 'fill.patternPreset',
			kind: 'select',
			label: t('pattern', 'Pattern'),
			value: pattern.preset,
			choices: [...presets],
		},
		{
			id: 'fill.patternForeground',
			kind: 'color',
			label: t('foreground', 'Foreground'),
			value: hexForInput(pattern.foreground),
		},
		{
			id: 'fill.patternBackground',
			kind: 'color',
			label: t('background', 'Background'),
			value: hexForInput(pattern.background, '#FFFFFF'),
		},
	];
}

/**
 * Office applies a texture when Picture or texture fill is chosen with no picture yet; this small
 * tiled parchment stands in for its default in every product.
 */
export const DEFAULT_CHART_TEXTURE =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAABoUlEQVR42j3S504CQRQF4Hn/19BojCUxMYiiomIBFKmCIh1sgL389sMxJpPZnXPvKXd2QyW70qlt37UOn0f5m/Jmq5L8HJcm/dN2dev25mA6yObSc+PeSa++E6vhvp2xvPUbu3atjmpozdIGGpyWhQwPHhYHBzU7zvB6l+3wOo3sJRpqgwQHYi+3Z93LFJSqAj0ZRIJoeL07Z/7n8DTMqT12j9vVzZjn/eFi1Nx/6BzpsKMRZt5vpIgG7F59Ow4QTY2I/z2toL3dF3DKp8t0NTAMylRRo7DEuuPccub35+GcaX08FqkEA+izWHGnhKAsAO14y/bB1R5cTyDjTI+7s0itSgIBrhUHAneN5IGBarOUiMMVMoscBCVGwmXIgyAnOXPW8msB9WtSbhTW+ZgEEyK0DCS18jGSLxPB0Kklncmz06Gbkr7I91I8XlICCsJ5NrTEatjcoZpI0KvmViVUjbdnEllm1wr1ackIKrpUo+aemgHscBeKz9wK8REHRTNWHFckO3L8ZZjQnf0av//P7B7PDxf+/2o1IA4HNEeiGij+AIB5gfnxhkAIAAAAAElFTkSuQmCC';

/** Picture source, transparency and tiling of a picture fill. */
export function pictureFields(picture: ChartFormatFillView['picture']): ChartFormatField[] {
	const fields: ChartFormatField[] = [
		{
			id: 'fill.picture',
			kind: 'file',
			label: t('pictureSource', 'Picture source'),
			accept: 'image/png,image/jpeg,image/gif,image/bmp',
			button: t('insertPicture', 'Insert...'),
			value: Boolean(picture?.src || picture?.relId),
		},
	];
	if (picture?.src || picture?.relId) {
		fields.push(
			{
				id: 'fill.pictureTransparency',
				kind: 'number',
				label: t('transparency', 'Transparency'),
				value: Math.round(picture.transparency ?? 0),
				min: 0,
				max: 100,
				step: 1,
				unit: '%',
			},
			{
				id: 'fill.pictureTile',
				kind: 'checkbox',
				label: t('pictureTile', 'Tile picture as texture'),
				value: Boolean(picture.tile),
			},
		);
	}
	return fields;
}
