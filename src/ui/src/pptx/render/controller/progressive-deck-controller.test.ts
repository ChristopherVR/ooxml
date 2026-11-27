import { describe, expect, it } from 'vitest';
import type { PptxSlide } from 'ooxml-core/pptx';

import {
	createProgressiveDeckController,
	decideProgressiveOperation,
	fillPendingSlides,
	isPendingSlide,
	pendingSlide,
	progressiveSearchScope,
	progressiveShowStep,
} from './progressive-deck-controller';

const parsed = (index: number, text = ''): PptxSlide => ({
	id: `slide${index + 1}`,
	rId: `rId${index + 2}`,
	slideNumber: index + 1,
	elements: [],
	notes: text,
});

describe('progressive deck rules', () => {
	it('fills only placeholders, wherever the user moved them', () => {
		const edited = { ...parsed(0), notes: 'edited' };
		// The user moved the third (pending) slide to the front and deleted the second.
		const slides = [pendingSlide(2), edited];
		const arrived = new Map([
			[0, parsed(0, 'stale parse')],
			[1, parsed(1)],
			[2, parsed(2)],
		]);
		const filled = fillPendingSlides(slides, arrived);
		expect(filled.map((slide) => slide.id)).toEqual(['slide3', 'slide1']);
		expect(filled[1]).toBe(edited);
		expect(fillPendingSlides(filled, arrived)).toBe(filled);
	});

	it('fills an undo snapshot taken before a slide arrived', () => {
		const controller = createProgressiveDeckController(2);
		const snapshot = [parsed(0), pendingSlide(1)];
		controller.arrive(1, parsed(1));
		expect(fillPendingSlides(snapshot, controller.arrived).every((s) => !isPendingSlide(s))).toBe(
			true,
		);
	});

	it('runs edits on parsed slides, blocks placeholders and queues deck work', () => {
		expect(decideProgressiveOperation({ kind: 'editSlide', slide: parsed(0) }, false)).toBe('run');
		expect(decideProgressiveOperation({ kind: 'editSlide', slide: pendingSlide(1) }, false)).toBe(
			'block',
		);
		expect(decideProgressiveOperation({ kind: 'arrangeSlides' }, false)).toBe('run');
		expect(
			decideProgressiveOperation(
				{ kind: 'copySlides', slides: [parsed(0), pendingSlide(1)] },
				false,
			),
		).toBe('queue');
		expect(decideProgressiveOperation({ kind: 'deck' }, false)).toBe('queue');
		expect(decideProgressiveOperation({ kind: 'editSlide', slide: pendingSlide(1) }, true)).toBe(
			'run',
		);
	});

	it('runs queued whole-deck work in order once the parse completes', async () => {
		const controller = createProgressiveDeckController(2);
		const order: string[] = [];
		const save = controller.whenComplete(() => order.push('save'));
		const replace = controller.whenComplete(() => order.push('replace'));
		expect(controller.getSnapshot()).toMatchObject({ queued: 2, complete: false });
		controller.arrive(0, parsed(0));
		controller.arrive(1, parsed(1));
		expect(order).toEqual([]);
		controller.finish();
		await Promise.all([save, replace]);
		expect(order).toEqual(['save', 'replace']);
		expect(controller.getSnapshot()).toMatchObject({ parsed: 2, complete: true, queued: 0 });
	});

	it('drops queued work when the load fails', async () => {
		const controller = createProgressiveDeckController(3);
		let ran = false;
		const save = controller.whenComplete(() => {
			ran = true;
		});
		controller.fail();
		await expect(save).rejects.toThrow();
		expect(ran).toBe(false);
		controller.arrive(0, parsed(0));
		expect(controller.getSnapshot().parsed).toBe(0);
	});

	it('searches parsed slides with their coverage and steps a show up to a placeholder', () => {
		const slides = [parsed(0), pendingSlide(1), parsed(2)];
		const scope = progressiveSearchScope(slides);
		expect(scope.slides.map(({ index }) => index)).toEqual([0, 2]);
		expect(scope).toMatchObject({ searched: 2, total: 3 });
		expect(progressiveShowStep(slides, 0)).toBe('ready');
		expect(progressiveShowStep(slides, 1)).toBe('wait');
		expect(progressiveShowStep(slides, 3)).toBe('end');
	});
});
