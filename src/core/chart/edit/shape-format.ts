// The chart Format ribbon tab's edits (Shape Styles, Shape Fill, Shape Outline, Shape Effects,
// WordArt Styles and Reset to Match Style) as typed `ChartEditIntent`s on one chart element, so
// every product's tab applies them through the same `applyChartEditIntent` its Format pane uses.
import type { DrawingOuterShadow } from '../../drawingml/effect-model';
import { defaultOuterShadow } from '../../drawingml/shape-effects-catalog';
import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import { chartTargetEffects } from './effects';
import type { ChartEditIntent } from './intent';
import type { ChartFormatTarget } from './target';

/** A resolved shape style: the fill, outline and outer shadow (null removes it) to apply. */
export interface ChartShapeFormat {
	fill?: DrawingFill;
	line?: DrawingLine;
	shadow?: DrawingOuterShadow | null;
}

/** A Shape Styles pick (or any resolved style) on `target`. */
export function chartShapeFormatIntents(
	target: ChartFormatTarget,
	format: ChartShapeFormat,
): ChartEditIntent[] {
	const intents: ChartEditIntent[] = [];
	if (format.fill) intents.push({ op: 'fill', target, fill: format.fill });
	if (format.line) intents.push({ op: 'line', target, line: format.line });
	if (format.shadow !== undefined)
		intents.push(chartShadowIntent(target, !!format.shadow, format.shadow));
	return intents;
}

/** Shape Fill: a colour, or none. */
export const chartFillIntent = (
	target: ChartFormatTarget,
	color: DrawingColor | null,
): ChartEditIntent => ({
	op: 'fill',
	target,
	fill: color ? { kind: 'solid', color } : { kind: 'none' },
});

/** Shape Outline: a colour (keeping the width and dash), or no line. */
export const chartOutlineIntent = (
	target: ChartFormatTarget,
	color: DrawingColor | null,
): ChartEditIntent => ({
	op: 'line',
	target,
	line: { fill: color ? { kind: 'solid', color } : { kind: 'none' } },
	merge: true,
});

/** Shape Effects > Shadow: the default outer shadow, or none. */
export function chartShadowIntent(
	target: ChartFormatTarget,
	enabled: boolean,
	shadow?: DrawingOuterShadow | null,
): ChartEditIntent {
	return {
		op: 'effect',
		target,
		edit: enabled ? (shadow ?? defaultOuterShadow()) : { kind: 'outerShadow', remove: true },
	};
}

/**
 * Reset to Match Style: the element's fill and outline return to the chart style, and the effects
 * it carries are removed (an element without its own effects is left as it is).
 */
export function chartResetFormatIntents(
	space: ChartSpace,
	target: ChartFormatTarget,
): ChartEditIntent[] {
	const kinds = new Set(
		(chartTargetEffects(space, target)?.effects ?? []).flatMap((effect) =>
			effect.kind === 'unsupported' ? [] : [effect.kind],
		),
	);
	return [
		{ op: 'fill', target, fill: undefined },
		{ op: 'line', target, line: undefined },
		...[...kinds].map((kind): ChartEditIntent => ({
			op: 'effect',
			target,
			edit: { kind, remove: true },
		})),
	];
}

/** A WordArt style, or Text Fill / Text Outline, on the element's text (null clears an outline). */
export function chartTextFormatIntents(
	target: ChartFormatTarget,
	style: { fill?: DrawingColor; outline?: DrawingColor | null; shadow?: boolean },
): ChartEditIntent[] {
	const intents: ChartEditIntent[] = [];
	const edit: { fill?: DrawingFill; line?: DrawingLine | null } = {};
	if (style.fill) edit.fill = { kind: 'solid', color: style.fill };
	if (style.outline !== undefined)
		edit.line = style.outline
			? { widthEmu: 9525, fill: { kind: 'solid', color: style.outline } }
			: null;
	if (edit.fill || edit.line !== undefined) intents.push({ op: 'text', target, edit });
	if (style.shadow !== undefined)
		intents.push({
			op: 'text',
			target,
			edit: {
				effect: style.shadow ? defaultOuterShadow() : { kind: 'outerShadow', remove: true },
			},
		});
	return intents;
}
