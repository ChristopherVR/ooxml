// @vitest-environment jsdom
import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createDefaultRegistry, renderSlideStage } from '../render';
import { playTransitionOverlay } from './transition-overlay';

/**
 * Issue #141, slides 8 -> 9: a full-slide picture fades in through its
 * transparency (`a:alphaModFix`). The plan tweens it with a rule on the
 * picture's `<img>` (`[data-element-id="..."] img`), a DESCENDANT selector:
 * it only reaches the picture if the stage this binding renders puts the
 * `<img>` inside the node carrying `data-element-id`. Rendered through the
 * real `renderSlideStage`, since the overlay layers are clones of it.
 */
const PIXEL = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

function picture(id: string, alphaModFix?: number, overrides: Partial<PptxElement> = {}) {
	return {
		id,
		name: id,
		type: 'picture',
		x: 0,
		y: 0,
		width: 960,
		height: 540,
		imagePath: 'ppt/media/faded.png',
		imageData: PIXEL,
		...(alphaModFix === undefined ? {} : { imageEffects: { alphaModFix } }),
		...overrides,
	} as PptxElement;
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, rId: `r-${id}`, slideNumber: 1, elements } as unknown as PptxSlide;
}

function stageOf(model: PptxSlide): HTMLElement {
	return renderSlideStage({
		document,
		slide: model,
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator('en'),
		scale: 1,
		interactive: false,
	});
}

describe('vanilla: a picture faded in through its transparency', () => {
	let stageWrap: HTMLElement;
	/** Every overlay a test started, cancelled even when an assertion throws first. */
	const overlays: (() => void)[] = [];

	beforeEach(() => {
		vi.useFakeTimers();
		stageWrap = document.createElement('div');
		document.body.appendChild(stageWrap);
	});

	afterEach(() => {
		for (const cancel of overlays.splice(0)) {
			cancel();
		}
		vi.useRealTimers();
		stageWrap.remove();
	});

	it('paints the <img> the alpha keyframe rule targets, on the stage copy and the ghost', () => {
		// The band dissolving out underneath makes the overlay paint the faded
		// picture's ghost too (see `resolveMorphGhostIds`).
		const from = slide('s8', [
			picture('s8-band', undefined, { imagePath: 'ppt/media/band.png', y: 340, height: 200 }),
			picture('s8-faded', 0),
		]);
		const to = slide('s9', [picture('s9-faded', 100)]);
		const cancel = playTransitionOverlay({
			doc: document,
			stageWrap,
			outgoing: stageOf(from),
			incoming: stageOf(to),
			transition: { type: 'morph', durationMs: 2000 } as PptxSlideTransition,
			outgoingSlide: from,
			incomingSlide: to,
			onDone: vi.fn(),
		});
		overlays.push(cancel);

		const css = [...stageWrap.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		for (const [scope, id] of [
			['data-pptx-morph-incoming', 's9-faded'],
			['data-pptx-morph-outgoing', 's8-faded'],
		] as const) {
			const selector = `[${scope}] [data-element-id="${id}"] img`;
			// The plan tweens the picture's transparency on its <img>...
			expect(css, `${selector} must carry the alpha tween`).toContain(`${selector} { animation:`);
			// ...and the overlay really paints an <img> that selector matches.
			const img = stageWrap.querySelector<HTMLImageElement>(selector);
			expect(img, `${selector} must match a painted <img>`).not.toBeNull();
			expect(img!.closest(`[data-element-id="${id}"]`)).not.toBeNull();
		}
		cancel();
	});
});
