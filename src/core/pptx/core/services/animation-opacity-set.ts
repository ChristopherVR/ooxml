import type { PptxNativeAnimation } from '../types';

/** A fixed opacity value with its own active interval, not an interpolated fade. */
export interface PptxOpacitySet {
	value: number;
	delayMs: number;
	durationMs: number;
	holdEndState: boolean;
}

const holds = (fill: string | undefined) => fill === 'hold' || fill === 'freeze';

/**
 * Resolve one isolated numeric opacity set. Native PowerPoint also emits a
 * legacy image filter hint, which does not turn this discrete set into a tween.
 * Missing fill restores after the active interval, as measured in native video.
 * Repeated, additive, reversed and mixed behavior trees remain unsupported here.
 */
export function resolveOpacitySet(anim: PptxNativeAnimation): PptxOpacitySet | undefined {
	if (anim.afterAnimationAction !== undefined) return undefined;
	if (
		anim.target?.type === 'shape' &&
		(anim.target.backgroundOnly ||
			anim.target.subShapeId ||
			anim.target.graphicElement ||
			anim.target.oleChartElement)
	)
		return undefined;
	if (
		anim.textTarget ||
		(anim.target && anim.target.type !== 'shape') ||
		anim.trigger === 'onHover'
	)
		return undefined;
	if (
		anim.presetClass !== 'emph' ||
		anim.autoReverse ||
		(anim.repeatCount ?? 1) !== 1 ||
		anim.repeatDurMs !== undefined
	)
		return undefined;
	const behaviors = anim.behaviors;
	if (
		!behaviors?.length ||
		behaviors.some(
			(b) =>
				b.kind !== 'set' &&
				!(
					b.kind === 'animEffect' &&
					b.filter === 'image' &&
					(b.transition === undefined || b.transition === 'none')
				),
		)
	)
		return undefined;
	const sets = behaviors.filter((b) => b.kind === 'set');
	if (sets.length !== 1) return undefined;
	const set = sets[0]!;
	if (
		set.attrNames.length !== 1 ||
		set.attrNames[0] !== 'style.opacity' ||
		(set.additive !== undefined && !['base', 'repl', 'none'].includes(set.additive))
	)
		return undefined;
	const raw = set.value;
	if (
		typeof raw !== 'number' &&
		(typeof raw !== 'string' || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(raw.trim()))
	)
		return undefined;
	const value = Number(raw);
	const { timing } = set;
	if (timing.unsupportedSetTiming) return undefined;
	if (anim.fill === 'transition' || timing.fill === 'transition') return undefined;
	if (
		!Number.isFinite(value) ||
		value < 0 ||
		value > 1 ||
		timing.autoReverse ||
		(timing.repeatCount ?? 1) !== 1 ||
		timing.tmFilter !== undefined
	)
		return undefined;
	const delayMs = timing.delayMs ?? 0;
	const durationMs = timing.durationMs;
	if (
		!Number.isFinite(delayMs) ||
		delayMs < 0 ||
		durationMs === undefined ||
		!Number.isFinite(durationMs) ||
		durationMs < 0
	)
		return undefined;
	const effectDuration = anim.durationMs ?? delayMs + durationMs;
	if (
		!Number.isFinite(effectDuration) ||
		effectDuration < delayMs ||
		(anim.speedPct ?? 100) !== 100
	)
		return undefined;
	const holdEndState = holds(anim.fill) && holds(timing.fill);
	return {
		value,
		delayMs,
		durationMs: holds(timing.fill)
			? effectDuration - delayMs
			: Math.min(durationMs, effectDuration - delayMs),
		holdEndState,
	};
}
