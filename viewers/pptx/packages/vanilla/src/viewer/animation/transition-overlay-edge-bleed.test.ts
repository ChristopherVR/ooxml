import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createDefaultRegistry, renderSlideStage } from '../render';
import { playTransitionOverlay } from './transition-overlay';

/**
 * The 2 px an abutting rectangle grows under its cover (abutting-edge-bleed)
 * is only hidden while the cover stays on top. This binding's morph overlay
 * clones the stage it was rendered into, so a clone of the grown rectangle
 * moved apart from its cover would show the growth as a sliver.
 */

function rect(id: string, x: number, color: string): PptxElement {
	return {
		type: 'shape',
		id,
		name: `!!${id.split('-')[1]}`,
		x,
		y: 100,
		width: 300,
		height: 300,
		shapeType: 'rect',
		text: '',
		shapeStyle: { fillMode: 'solid', fillColor: color, strokeWidth: 0, strokeFillMode: 'none' },
	} as unknown as PptxElement;
}

const slideA = {
	id: 'a',
	elements: [rect('a-lower', 300, '#C81E1E'), rect('a-upper', 600, '#1414A0')],
} as unknown as PptxSlide;
const slideB = {
	id: 'b',
	elements: [rect('b-lower', 50, '#C81E1E'), rect('b-upper', 900, '#1414A0')],
	transition: { type: 'morph', durationMs: 800 },
} as unknown as PptxSlide;

function stageOf(slide: PptxSlide): HTMLElement {
	return renderSlideStage({
		document,
		slide,
		canvasSize: { width: 1280, height: 720 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: ((key: string) => key) as never,
	});
}

describe('morph overlay and the abutting-edge bleed', () => {
	let stageWrap: HTMLElement;
	beforeEach(() => {
		vi.useFakeTimers();
		stageWrap = document.createElement('div');
		document.body.appendChild(stageWrap);
	});
	afterEach(() => {
		vi.useRealTimers();
		stageWrap.remove();
	});

	it('renders the outgoing stage with the lower rectangle grown (the premise)', () => {
		const lower = stageOf(slideA).querySelector<HTMLElement>('[data-element-id="a-lower"]')!;
		expect(lower.style.width).toBe('302px');
	});

	it('paints no grown rectangle in any copy the morph overlay makes', () => {
		const outgoing = stageOf(slideA);
		const cancel = playTransitionOverlay({
			doc: document,
			stageWrap,
			outgoing,
			incoming: stageOf(slideB),
			transition: slideB.transition as PptxSlideTransition,
			outgoingSlide: slideA,
			incomingSlide: slideB,
			onDone: vi.fn(),
		});
		const copies = [...stageWrap.querySelectorAll<HTMLElement>('[data-element-id]')];
		expect(copies.length).toBeGreaterThan(0);
		for (const copy of copies) {
			expect(copy.style.width).toBe('300px');
		}
		expect(stageWrap.querySelector('[data-pptx-edge-bleed]')).toBeNull();
		cancel();
	});
});
