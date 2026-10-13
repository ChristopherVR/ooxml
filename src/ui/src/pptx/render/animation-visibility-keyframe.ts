import type { PptxNativeAnimation } from 'ooxml-core/pptx';
import { resolveVisibilitySetTiming } from 'ooxml-core/pptx';

/** Paint a held native visibility flip at its child offset, without interpolated opacity. */
export function buildVisibilitySetKeyframe(
	anim: PptxNativeAnimation,
	uid: number,
):
	| {
			keyframeName: string;
			css: string;
			easing: string;
	  }
	| undefined {
	const timing = resolveVisibilitySetTiming(anim);
	if (!timing) return undefined;
	const duration = timing.delayMs + timing.durationMs;
	if (!Number.isFinite(duration) || duration <= 0) return undefined;
	const keyframeName = `pptx-tl-visibility-${uid}`;
	const change = (timing.delayMs / duration) * 100;
	const frames =
		timing.delayMs === 0
			? '0%,100%{visibility:visible;}'
			: `0%{visibility:hidden;}${change}%{visibility:visible;}100%{visibility:visible;}`;
	return { keyframeName, css: `@keyframes ${keyframeName}{${frames}}`, easing: 'steps(1,end)' };
}
