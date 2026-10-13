// @vitest-environment jsdom
import type { PptxNativeAnimation } from 'ooxml-core/pptx';
import { expect, it, vi } from 'vitest';
import { buildTimeline } from './animation-timeline-builder';
import { TimelineEngine } from './animation-timeline-engine';
import { applyAnimationGroupSteps } from './animation-playback-engine';
import { clearPlaybackTimers } from './animation-playback-seek';
import type { PlaybackContext } from './animation-playback-engine';
import type { ElementAnimationState } from './animation-timeline-types';
import { keyframeAnimatedProperties } from './animation-parallel-composition';

const native = (delay: number): PptxNativeAnimation => ({
	targetId: 's',
	trigger: 'onClick',
	presetClass: 'entr',
	presetId: 0,
	durationMs: delay + 1,
	visibilitySetTiming: { delayMs: delay, durationMs: 1 },
	seqNextAction: 'seek',
	seqConcurrent: true,
	fill: 'hold',
	target: { type: 'shape', shapeId: '2' },
	behaviors: [
		{
			kind: 'set',
			attrNames: ['style.visibility'],
			value: 'visible',
			timing: { delayMs: delay, durationMs: 1, fill: 'hold' },
		},
	],
});
it.each([0, 499, 1499])('keeps the native %ims child offset without a fade', (delay) => {
	const timeline = buildTimeline([native(delay)]);
	expect(timeline.clickGroups[0]?.totalDurationMs).toBe(delay + 1);
	expect(timeline.clickGroups[0]?.steps[0]?.cssAnimation).toContain('steps(1,end)');
	expect(timeline.keyframesCss).toContain('visibility:visible');
	expect(timeline.keyframesCss).not.toContain('opacity:');
	expect(timeline.keyframesCss).not.toContain('pptx-fadeIn');
	if (delay) expect(timeline.keyframesCss).toContain(`${(delay / (delay + 1)) * 100}%`);
});
it('projects the held end state through completion and resets to hidden on restart', () => {
	const engine = new TimelineEngine(buildTimeline([native(1499)]));
	expect(engine.getElementStates(['s']).get('s')?.visible).toBe(false);
	engine.completeAll();
	expect(engine.getElementStates(['s']).get('s')?.visible).toBe(true);
	engine.reset();
	expect(engine.getElementStates(['s']).get('s')?.visible).toBe(false);
});
it('keeps visibility and opacity CSS components independent during playback and cleanup', () => {
	vi.useFakeTimers();
	try {
		const group = buildTimeline([native(1499)]).clickGroups[0]!;
		const visibility = group.steps[0]!;
		const opacity = {
			...visibility,
			presetClass: 'emph' as const,
			keyframeName: 'pptx-tl-tav-opacity-test',
			cssAnimation: 'pptx-tl-tav-opacity-test 1500ms linear 0ms 1 normal both',
		};
		expect(keyframeAnimatedProperties(visibility.keyframeName)).toEqual(new Set(['visibility']));
		expect(keyframeAnimatedProperties(opacity.keyframeName)).toEqual(new Set(['opacity']));
		let states = new Map<string, ElementAnimationState>();
		const ctx: PlaybackContext = {
			timers: [],
			buildHandle: { current: null },
			setStates: (update) => {
				states = update(states);
			},
			playSound: () => undefined,
			stopSound: () => undefined,
		};
		applyAnimationGroupSteps({ ...group, steps: [visibility, opacity] }, ctx);
		expect(states.get('s')?.cssAnimation).toBe(
			`${visibility.cssAnimation}, ${opacity.cssAnimation}`,
		);
		clearPlaybackTimers(ctx);
		states = new Map([['s', { visible: false, cssAnimation: undefined }]]);
		vi.advanceTimersByTime(1600);
		expect(states.get('s')).toEqual({ visible: false, cssAnimation: undefined });
		expect(vi.getTimerCount()).toBe(0);
	} finally {
		vi.useRealTimers();
	}
});
it('cancels the longer native window when stopped and cannot change restarted state', () => {
	vi.useFakeTimers();
	try {
		let states = new Map<string, ElementAnimationState>([
			['s', { visible: false, cssAnimation: undefined }],
		]);
		const ctx: PlaybackContext = {
			timers: [],
			buildHandle: { current: null },
			setStates: (update) => {
				states = update(states);
			},
			playSound: () => undefined,
			stopSound: () => undefined,
		};
		applyAnimationGroupSteps(buildTimeline([native(1499)]).clickGroups[0]!, ctx);
		expect(ctx.timers.length).toBeGreaterThan(0);
		clearPlaybackTimers(ctx);
		states = new Map([['s', { visible: false, cssAnimation: undefined }]]);
		vi.advanceTimersByTime(1600);
		expect(states.get('s')?.visible).toBe(false);
		expect(ctx.timers).toEqual([]);
		expect(vi.getTimerCount()).toBe(0);
	} finally {
		vi.useRealTimers();
	}
});
it('reports unsupported preset-zero visibility instead of inventing a fade', () => {
	const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
	try {
		const timeline = buildTimeline([
			{
				targetId: 's',
				presetClass: 'entr',
				presetId: 0,
				behaviors: [
					{
						kind: 'set',
						attrNames: ['style.visibility'],
						value: 'visible',
						timing: { unsupportedSetTiming: true },
					},
				],
			},
		]);
		expect(timeline.clickGroups).toHaveLength(0);
		expect(warn).toHaveBeenCalledWith(
			'Unsupported native visibility set target or timing profile',
			's',
		);
	} finally {
		warn.mockRestore();
	}
});

it('reports a retained visibility body even when no value behavior was decoded', () => {
	const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
	try {
		const timeline = buildTimeline([
			{ targetId: 's', presetClass: 'entr', presetId: 0, unsupportedVisibilitySet: true },
		]);
		expect(timeline.clickGroups).toHaveLength(0);
		expect(warn).toHaveBeenCalledWith(
			'Unsupported native visibility set target or timing profile',
			's',
		);
	} finally {
		warn.mockRestore();
	}
});
