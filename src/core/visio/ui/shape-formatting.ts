import type { VisioShape } from '../model';
import {
	VISIO_SHADOW_PRESETS,
	visioShadowPresetGeometry,
	type VisioShadowPreset,
} from '../edit-formatting-effects';

export interface VisioShapeFormattingState {
	linePattern: number | undefined;
	fillPatternIndex: number | undefined;
	fillBackgroundColor: string | undefined;
	fillForegroundTransparency: number | undefined;
	fillBackgroundTransparency: number | undefined;
	/** Mixed foreground/background values are undefined, including within one shape. */
	fillTransparency: number | undefined;
	lineTransparency: number | undefined;
	/** The common outer shadow preset; undefined when mixed or not a preset. */
	shadowPreset: VisioShadowPreset | undefined;
}
/** The preset whose offset and blur match a rendered shadow, or `none` without one. */
export function visioShadowPreset(shape: VisioShape): VisioShadowPreset | undefined {
	const shadow = shape.style.shadow;
	if (!shadow) return 'none';
	const near = (left: number, right: number) => Math.abs(left - right) < 1e-6;
	return VISIO_SHADOW_PRESETS.find((preset) => {
		if (preset === 'none') return false;
		const { x, y, blur } = visioShadowPresetGeometry(preset);
		return (
			near(shadow.offsetX, x / 72) && near(shadow.offsetY, y / 72) && near(shadow.blur, blur / 72)
		);
	});
}
/** Source paint values stay independent of gradient and bitmap rendering multipliers. */
export function visioShapeFormattingState(
	shapes: readonly VisioShape[],
): VisioShapeFormattingState {
	const common = <T>(values: readonly T[]): T | undefined =>
		values.length && values.every((value) => value === values[0]) ? values[0] : undefined;
	const percent = (opacity: number | undefined): number | undefined =>
		opacity === undefined || !Number.isFinite(opacity) || opacity < 0 || opacity > 1
			? undefined
			: Math.round((1 - opacity) * 200) / 2;
	const foreground = shapes.map((shape) =>
		percent(
			shape.style.fillForegroundOpacity ??
				(shape.style.fillGradient || shape.style.fillPattern ? undefined : shape.style.fillOpacity),
		),
	);
	const background = shapes.map((shape) => percent(shape.style.fillBackgroundOpacity));
	return {
		linePattern: common(shapes.map((shape) => shape.style.linePattern)),
		fillPatternIndex: common(shapes.map((shape) => shape.style.fillPatternIndex)),
		fillBackgroundColor: common(shapes.map((shape) => shape.style.fillBackgroundColor)),
		fillForegroundTransparency: common(foreground),
		fillBackgroundTransparency: common(background),
		fillTransparency: common([...foreground, ...background]),
		shadowPreset: common(shapes.map(visioShadowPreset)),
		lineTransparency: common(
			shapes.map((shape) =>
				percent(
					shape.style.lineColorOpacity ??
						(shape.style.lineGradient ? undefined : shape.style.lineOpacity),
				),
			),
		),
	};
}
