import type { PlaybackContext } from './animation-playback-engine';
import type { ElementAnimationState, TimelineClickGroup } from './animation-timeline-types';

interface ActiveSet {
	value: number;
}
interface OpacityStack {
	baseline: number | undefined;
	entries: ActiveSet[];
}
const stacks = new WeakMap<number[], Map<string, OpacityStack>>();
const pending = new WeakMap<number[], Map<TimelineClickGroup['steps'][number], () => void>>();

/** Drop old held sets when the owning show/slide resets its timers. */
export function resetOpacitySetPlayback(timers: number[]): void {
	stacks.delete(timers);
	pending.delete(timers);
}

/** Seek settles even a delayed fixed set before cancelling its timers. */
export function finishOpacitySetSteps(group: TimelineClickGroup, ctx: PlaybackContext): void {
	for (const step of group.steps) pending.get(ctx.timers)?.get(step)?.();
}

/** Apply fixed values at their authored starts and restore the surviving prior set. */
export function applyOpacitySetSteps(group: TimelineClickGroup, ctx: PlaybackContext): void {
	const targets = stacks.get(ctx.timers) ?? new Map<string, OpacityStack>();
	stacks.set(ctx.timers, targets);
	const completions = pending.get(ctx.timers) ?? new Map();
	pending.set(ctx.timers, completions);
	for (const step of group.steps) {
		const set = step.opacitySet;
		if (!set || !step.elementId) continue;
		const entry: ActiveSet = { value: set.value };
		const startMs = step.delayMs + set.delayMs;
		let started = false;
		const start = () => {
			if (started) return;
			started = true;
			ctx.setStates((previous) => {
				const current = previous.get(step.elementId);
				const stack = targets.get(step.elementId) ?? { baseline: current?.opacity, entries: [] };
				stack.entries.push(entry);
				targets.set(step.elementId, stack);
				const next = new Map(previous);
				next.set(step.elementId, {
					...current,
					visible: current?.visible ?? true,
					cssAnimation: current?.cssAnimation,
					opacity: set.value,
				});
				return next;
			});
		};
		if (startMs > 0) ctx.timers.push(window.setTimeout(start, startMs));
		else start();
		const finish = () => {
			start();
			if (set.holdEndState) return;
			const stack = targets.get(step.elementId);
			if (!stack) return;
			const index = stack.entries.indexOf(entry);
			if (index < 0) return;
			stack.entries.splice(index, 1);
			const opacity = stack.entries.at(-1)?.value ?? stack.baseline;
			if (!stack.entries.length) targets.delete(step.elementId);
			ctx.setStates((previous) => {
				const current = previous.get(step.elementId);
				if (!current) return previous;
				const next = new Map(previous);
				const restored: ElementAnimationState = { ...current };
				if (opacity === undefined) delete restored.opacity;
				else restored.opacity = opacity;
				next.set(step.elementId, restored);
				return next;
			});
		};
		completions.set(step, finish);
		if (!set.holdEndState) ctx.timers.push(window.setTimeout(finish, startMs + set.durationMs));
	}
}
