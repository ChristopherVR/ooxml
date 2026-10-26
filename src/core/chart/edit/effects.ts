import { drawingEffectsXml, parseDrawingEffects } from '../../drawingml/drawing-effects';
import type { DrawingEffect, DrawingEffectList } from '../../drawingml/effect-model';
import { parseXml } from '../../xml/index';
import type { ChartSpace } from '../model';
import { editChartSpace } from './draft';
import type { ChartFormatTarget } from './target';
import { chartTargetDraftShapes, chartTargetShape } from './target-paint';
import { withDrawingColorTransparency } from './paint';

type KnownEffect = Exclude<DrawingEffect, { kind: 'unsupported' }>;
export type ChartEffectEdit =
	| KnownEffect
	| { kind: KnownEffect['kind']; remove: true }
	| { kind: 'outerShadow' | 'innerShadow' | 'presetShadow' | 'glow'; transparency: number };
const ORDER: readonly KnownEffect['kind'][] = [
	'blur',
	'fillOverlay',
	'glow',
	'innerShadow',
	'outerShadow',
	'presetShadow',
	'reflection',
	'softEdge',
];

export function chartTargetEffects(
	space: ChartSpace,
	target: ChartFormatTarget,
): DrawingEffectList | undefined {
	const xml = chartTargetShape(space, target)?.effectsXml;
	return xml ? parseDrawingEffects(parseXml(xml).documentElement) : undefined;
}

/** Applies one effect without clearing other effects, unknown nodes, or authored style tokens. */
export function editChartEffectList(
	list: DrawingEffectList | undefined,
	edit: ChartEffectEdit,
): DrawingEffectList {
	const next = structuredClone(list ?? { effects: [] });
	const matches = next.effects.flatMap((effect, index) =>
		effect.kind === edit.kind ? [index] : [],
	);
	if (matches.length > 1) throw new RangeError('Ambiguous duplicate chart effect');
	const index = matches[0];
	if ('transparency' in edit) {
		const effect = index === undefined ? undefined : next.effects[index];
		if (!effect || !('color' in effect) || !effect.color)
			throw new RangeError('Effect transparency needs an authored colour');
		effect.color = withDrawingColorTransparency(effect.color, edit.transparency);
		return next;
	}
	if ('remove' in edit) {
		if (index !== undefined) next.effects.splice(index, 1);
		return next;
	}
	if (index !== undefined)
		next.effects[index] = { ...next.effects[index], ...structuredClone(edit) } as DrawingEffect;
	else {
		const after = next.effects.findIndex(
			(effect) =>
				effect.kind !== 'unsupported' && ORDER.indexOf(effect.kind) > ORDER.indexOf(edit.kind),
		);
		next.effects.splice(after < 0 ? next.effects.length : after, 0, structuredClone(edit));
	}
	return next;
}

/** One target's shape effect; explicit empty lists suppress inherited theme effects. */
export function setChartTargetEffect(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartEffectEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target)) {
			const list = shape.effectsXml
				? parseDrawingEffects(parseXml(shape.effectsXml).documentElement)
				: undefined;
			if (shape.effectsXml && !list)
				throw new RangeError('The chart effect source is not an editable effect list');
			shape.effectsXml = drawingEffectsXml(editChartEffectList(list, edit));
		}
	});
}
