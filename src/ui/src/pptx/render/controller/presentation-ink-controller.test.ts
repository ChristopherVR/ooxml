import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	applyPresentationInk,
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
	presentationInkSlideCount,
} from './presentation-ink-controller';
import type { PresentationInkTimers } from './presentation-ink-controller';

function draw(
	controller: ReturnType<typeof createPresentationInkController>,
	points: [number, number][],
) {
	const [first, ...rest] = points;
	controller.pointerDown({ x: first![0], y: first![1] });
	for (const [x, y] of rest) {
		controller.pointerMove({ x, y });
	}
	controller.pointerUp();
}

describe('createPresentationInkController', () => {
	it('toggles the armed tool and keeps separate pen and highlighter colours', () => {
		const ink = createPresentationInkController();
		ink.toggleTool('pen');
		expect(ink.getSnapshot().tool).toBe('pen');
		ink.toggleTool('pen');
		expect(ink.getSnapshot().tool).toBe('none');
		ink.setTool('highlighter');
		ink.setTool('highlighter');
		expect(ink.getSnapshot().tool).toBe('highlighter');
		ink.setHighlighterColor('#00ff00');
		draw(ink, [
			[0, 0],
			[10, 10],
		]);
		expect(ink.getSnapshot().strokes[0]).toMatchObject({
			tool: 'highlighter',
			color: '#00ff00',
			width: 14,
			opacity: 0.4,
		});
		expect(ink.getSnapshot().penColor).toBe('#ff0000');
	});

	it('keeps strokes per slide and drops single-point strokes', () => {
		const ink = createPresentationInkController();
		ink.setTool('pen');
		draw(ink, [[5, 5]]);
		expect(ink.getSnapshot().strokes).toHaveLength(0);
		draw(ink, [
			[0, 0],
			[1, 1],
		]);
		ink.setSlide(3);
		draw(ink, [
			[2, 2],
			[3, 3],
		]);
		const state = ink.getSnapshot();
		expect(presentationInkOnSlide(state)).toHaveLength(1);
		expect(presentationInkOnSlide(state, 0)).toHaveLength(1);
		expect(presentationInkSlideCount(state)).toBe(2);
		expect([...presentationInkBySlide(state).keys()]).toEqual([0, 3]);
	});

	it('erases only on the current slide, within the radius', () => {
		const ink = createPresentationInkController({ eraserRadius: 5 });
		ink.setTool('pen');
		draw(ink, [
			[0, 0],
			[1, 0],
		]);
		ink.setSlide(1);
		draw(ink, [
			[0, 0],
			[1, 0],
		]);
		ink.setTool('eraser');
		ink.pointerMove({ x: 0, y: 0 });
		expect(ink.getSnapshot().strokes).toHaveLength(2);
		ink.pointerDown({ x: 20, y: 20 });
		ink.pointerMove({ x: 2, y: 2 });
		ink.pointerUp();
		expect(ink.getSnapshot().strokes.map((s) => s.slideIndex)).toEqual([0]);
	});

	it('moves and hides the laser; clears a slide or all', () => {
		const ink = createPresentationInkController();
		ink.setTool('laser');
		ink.pointerMove({ x: 4, y: 5 });
		expect(ink.getSnapshot().laser).toEqual({ x: 4, y: 5 });
		ink.pointerLeave();
		expect(ink.getSnapshot().laser).toBeNull();
		ink.setTool('pen');
		draw(ink, [
			[0, 0],
			[1, 1],
		]);
		ink.setSlide(1);
		draw(ink, [
			[0, 0],
			[1, 1],
		]);
		ink.clearSlide();
		expect(ink.getSnapshot().strokes).toHaveLength(1);
		ink.clearAll();
		expect(ink.getSnapshot().strokes).toHaveLength(0);
	});

	it('reset disarms but keeps the strokes for the exit prompt', () => {
		const ink = createPresentationInkController();
		ink.setTool('pen');
		draw(ink, [
			[0, 0],
			[1, 1],
		]);
		ink.setMarkupVisible(false);
		ink.reset();
		expect(ink.getSnapshot()).toMatchObject({ tool: 'none', laser: null, markupVisible: false });
		expect(ink.getSnapshot().strokes).toHaveLength(1);
	});

	it('hides the revealed toolbar after the delay', () => {
		let pending: (() => void) | null = null;
		const timers: PresentationInkTimers = {
			setTimeout: (handler) => {
				pending = handler;
				return 1;
			},
			clearTimeout: () => {
				pending = null;
			},
		};
		const ink = createPresentationInkController({ timers });
		ink.revealToolbar();
		expect(ink.getSnapshot().toolbarVisible).toBe(true);
		pending!();
		expect(ink.getSnapshot().toolbarVisible).toBe(false);
	});

	it('adopts session strokes with an opacity from their tool', () => {
		const ink = createPresentationInkController();
		ink.setStrokes([
			{ id: 'a', slideIndex: 0, tool: 'highlighter', color: '#ff0', width: 14, points: [] },
		]);
		expect(ink.getSnapshot().strokes[0]!.opacity).toBe(0.4);
	});
});

describe('applyPresentationInk', () => {
	const slides = [{ elements: [] }, { elements: [] }] as unknown as PptxSlide[];

	it('appends ink elements to their slides, scaled', () => {
		const next = applyPresentationInk(
			slides,
			[
				{
					id: 'a',
					slideIndex: 1,
					tool: 'pen',
					color: '#ff0000',
					width: 2,
					points: [
						{ x: 0.1, y: 0.1 },
						{ x: 0.5, y: 0.5 },
					],
				},
			],
			{ width: 960, height: 540 },
		);
		expect(next[0]).toBe(slides[0]);
		expect(next[1]!.elements).toHaveLength(1);
	});

	it('returns the slides unchanged when there is nothing to add', () => {
		expect(applyPresentationInk(slides, [])).toBe(slides);
	});
});
