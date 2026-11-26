import type { PptxElementAnimation } from 'ooxml-core/pptx';
import { expect, it, vi } from 'vitest';

import { applyAnimationPaint, captureAnimationPaint } from './animation-painter';
import { createAnimationPainterController } from './controller/animation-painter-controller';
import { editRibbonAnimationTiming } from './ribbon-animation-editing';

const source: PptxElementAnimation = {
	elementId: 'source',
	entrance: 'flyIn',
	emphasis: 'spin',
	exit: 'fadeOut',
	motionPath: 'M 0 0 L 1 1 E',
	direction: 'fromLeft',
	durationMs: 1250,
	delayMs: 400,
	timingCurve: 'ease-in',
	repeatCount: 2,
	trigger: 'onMediaBookmark',
	triggerShapeId: 'media',
	triggerBookmark: 'intro',
	order: 4,
	orderFromTimeline: 4,
};

it('replaces all editor effects on a target, keeping its order and dropping interactive triggers', () => {
	const target: PptxElementAnimation = { elementId: 'target', entrance: 'appear', order: 1 };
	const slide = { id: 'slide', animations: [source, target] };
	const paint = captureAnimationPaint(slide, source.elementId)!;
	expect(applyAnimationPaint(paint, slide, source.elementId)).toBeUndefined();
	const next = applyAnimationPaint(paint, slide, target.elementId)!;
	expect(next[0]).toBe(source);
	expect(next[1]).toStrictEqual({
		elementId: 'target',
		entrance: 'flyIn',
		emphasis: 'spin',
		exit: 'fadeOut',
		motionPath: 'M 0 0 L 1 1 E',
		direction: 'fromLeft',
		durationMs: 1250,
		delayMs: 400,
		timingCurve: 'ease-in',
		repeatCount: 2,
		trigger: 'onClick',
		order: 1,
	});
	expect(slide.animations[1]).toBe(target);
});

it('retains ordinary Start modes and appends after native timeline anchors', () => {
	const slide = {
		id: 'slide',
		animations: [{ ...source, trigger: 'withPrevious' } satisfies PptxElementAnimation],
		animationTimelineAnchors: [{ order: 7, targetIds: ['native'], presetClasses: [] }],
	};
	const paint = captureAnimationPaint(slide, 'source')!;
	expect(applyAnimationPaint(paint, slide, 'new')?.[1]).toMatchObject({
		elementId: 'new',
		order: 8,
		trigger: 'withPrevious',
	});
	expect(captureAnimationPaint(slide, 'native')).toBeUndefined();
});

it('keeps embedded sound on the source slide and clears its relationship on another slide', () => {
	const sound = { soundRId: 'rId5', soundPath: 'ppt/media/sound.wav', soundName: 'sound.wav' };
	const slide = { id: 'slide', animations: [{ ...source, ...sound }] };
	const paint = captureAnimationPaint(slide, 'source')!;
	expect(applyAnimationPaint(paint, slide, 'new')?.[1]).toMatchObject(sound);
	const [copy] = applyAnimationPaint(paint, { id: 'other-slide' }, 'source')!;
	expect(copy).toMatchObject({
		elementId: 'source',
		entrance: 'flyIn',
		soundPath: sound.soundPath,
		soundName: sound.soundName,
	});
	expect(copy.soundRId).toBeUndefined();
});

it('uses one commit per application and keeps the armed source in locked mode until cancelled', () => {
	const controller = createAnimationPainterController();
	const commit = vi.fn();
	const editing = { slide: { id: 'slide', animations: [source] }, elementId: 'source', commit };
	controller.arm(editing, 'once');
	controller.apply(editing, 'source');
	controller.apply(editing, 'first');
	controller.apply(editing, 'second');
	expect(commit).toHaveBeenCalledOnce();
	expect(controller.getSnapshot().mode).toBe('off');

	controller.arm(editing, 'once');
	controller.arm(editing, 'locked');
	const withoutSelection = { ...editing, elementId: undefined };
	controller.apply(withoutSelection, 'second');
	controller.apply(withoutSelection, 'third');
	controller.cancel();
	controller.apply(withoutSelection, 'fourth');
	expect(commit).toHaveBeenCalledTimes(3);
	expect(commit.mock.calls[2]?.[0][1]).toMatchObject({ elementId: 'third', entrance: 'flyIn' });
});

it('edits timing through the pane setters and skips no-ops', () => {
	const commit = vi.fn();
	const editing = { slide: { animations: [source] }, elementId: 'source', commit };
	editRibbonAnimationTiming(editing, { field: 'duration', value: Number.NaN });
	editRibbonAnimationTiming(editing, { field: 'delay', value: 400 });
	editRibbonAnimationTiming(editing, { field: 'move', value: 1 });
	expect(commit).not.toHaveBeenCalled();

	editRibbonAnimationTiming(editing, { field: 'start', value: 'withPrevious' });
	expect(commit).toHaveBeenCalledExactlyOnceWith([
		{ ...source, trigger: 'withPrevious', triggerShapeId: undefined, triggerBookmark: undefined },
	]);
});
