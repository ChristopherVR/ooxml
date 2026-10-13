import type { PptxNativeAnimation } from 'ooxml-core/pptx';
import { expect, it, vi } from 'vitest';
import { buildTimeline } from './animation-timeline-builder';
import { TimelineEngine } from './animation-timeline-engine';

const opacity = (hold: boolean): PptxNativeAnimation => ({
	targetId: 's',
	presetClass: 'emph',
	trigger: 'onClick',
	presetId: 9,
	durationMs: 500,
	fill: hold ? 'hold' : 'remove',
	behaviors: [
		{
			kind: 'set',
			attrNames: ['style.opacity'],
			value: '0.4',
			timing: { durationMs: 500, fill: hold ? 'hold' : 'remove' },
		},
	],
});

it('represents a native set as a fixed state step rather than a transparency pulse', () => {
	const timeline = buildTimeline([opacity(true)]);
	expect(timeline.clickGroups[0]?.steps[0]?.cssAnimation).toBe('');
	expect(timeline.clickGroups[0]?.steps[0]?.opacitySet?.value).toBe(0.4);
	expect(timeline.keyframesCss).not.toContain('pptx-transparency');
});
it('keeps committed end-state projection through later unrelated effects and resets it', () => {
	const timeline = buildTimeline([
		opacity(true),
		{ targetId: 's', presetClass: 'emph', trigger: 'onClick', presetId: 8 },
	]);
	const engine = new TimelineEngine(timeline);
	engine.advance();
	engine.advance();
	expect(engine.getElementStates(['s']).get('s')?.opacity).toBe(0.4);
	engine.reset();
	expect(engine.getElementStates(['s']).get('s')?.opacity).toBeUndefined();
	engine.completeAll();
	expect(engine.getElementStates(['s']).get('s')?.opacity).toBe(0.4);
});
it('does not project a removed set as a committed value', () => {
	const engine = new TimelineEngine(buildTimeline([opacity(false)]));
	engine.completeAll();
	expect(engine.getElementStates(['s']).get('s')?.opacity).toBeUndefined();
});
it('projects delayed opacity only during its active interval', () => {
	const anim = opacity(false);
	anim.behaviors![0]!.timing.delayMs = 100;
	anim.durationMs = 600;
	const engine = new TimelineEngine(buildTimeline([anim]));
	engine.advance(1000);
	expect(engine.getElementStates(['s'], { elapsedMs: 99 }).get('s')?.opacity).toBeUndefined();
	expect(engine.getElementStates(['s'], { elapsedMs: 100 }).get('s')?.opacity).toBe(0.4);
	expect(engine.getElementStates(['s'], { elapsedMs: 600 }).get('s')?.opacity).toBeUndefined();
});
it('does not project a delayed held set before its start', () => {
	const anim = opacity(true);
	anim.behaviors![0]!.timing.delayMs = 100;
	anim.durationMs = 600;
	const engine = new TimelineEngine(buildTimeline([anim]));
	engine.advance(1000);
	expect(engine.getElementStates(['s'], { elapsedMs: 99 }).get('s')?.opacity).toBeUndefined();
	expect(engine.getElementStates(['s'], { elapsedMs: 600 }).get('s')?.opacity).toBe(0.4);
	engine.reset();
	expect(engine.getElementStates(['s'], { elapsedMs: 600 }).get('s')?.opacity).toBeUndefined();
});
it('refuses and reports an unverified background-only numeric set', () => {
	const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
	try {
		const anim = opacity(true);
		anim.target = { type: 'shape', shapeId: '2', backgroundOnly: true };
		expect(buildTimeline([anim]).clickGroups).toEqual([]);
		expect(warning).toHaveBeenCalledWith(
			'Unsupported native opacity set target, value or timing profile',
			's',
		);
	} finally {
		warning.mockRestore();
	}
});
