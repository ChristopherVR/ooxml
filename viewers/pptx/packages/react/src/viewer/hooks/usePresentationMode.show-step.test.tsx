// @vitest-environment happy-dom
/**
 * The shared show step (`resolveShowForward` / `resolveShowBackward`) wired
 * into React's slide navigation: a forward press skips a hidden slide, past
 * the last slide the black end screen rises (the default "End with black
 * slide"), a backward press dismisses it, a click is gated on a slide whose
 * transition forbids click-advance while the keyboard is not, and a back press
 * at the start of the show stays put.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import type { ViewerMode } from '../types-core';
import { usePresentationMode } from './usePresentationMode';
import type { UsePresentationModeResult } from './usePresentationMode';

let container: HTMLDivElement;
let root: Root;
let latest: UsePresentationModeResult | null = null;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	latest = null;
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

/** Four slides; slide 2 (index 1) is hidden, slide 3 forbids click-advance. */
const slides: PptxSlide[] = [
	{ id: 's1', rId: 'rId1', elements: [] },
	{ id: 's2', rId: 'rId2', elements: [], hidden: true },
	{ id: 's3', rId: 'rId3', elements: [], transition: { type: 'none', advanceOnClick: false } },
	{ id: 's4', rId: 'rId4', elements: [] },
] as unknown as PptxSlide[];

function Harness({ mode }: { mode: ViewerMode }): React.ReactElement {
	const containerRef = useRef<HTMLDivElement | null>(null);
	latest = usePresentationMode({
		mode,
		slides,
		visibleSlideIndexes: [0, 2, 3],
		activeSlideIndex: 0,
		containerRef,
		onSetMode: () => {},
		onSetActiveSlideIndex: () => {},
	});
	return <div ref={containerRef} />;
}

function present(): UsePresentationModeResult {
	act(() => root.render(<Harness mode='edit' />));
	act(() => root.render(<Harness mode='present' />));
	return latest!;
}

const move = (direction: 1 | -1, trigger?: 'click' | 'explicit') =>
	act(() => latest!.movePresentationSlide(direction, trigger));

describe('react show step', () => {
	it('skips a hidden slide, raises the end screen past the last, and dismisses it backward', () => {
		present();
		move(1);
		expect(latest!.presentationSlideIndex).toBe(2);
		move(1);
		expect(latest!.presentationSlideIndex).toBe(3);
		move(1);
		expect(latest!.endOfShowVisible).toBeTruthy();
		move(-1);
		expect(latest!.endOfShowVisible).toBeFalsy();
		expect(latest!.presentationSlideIndex).toBe(3);
	});

	it('gates a click on a slide that forbids click-advance, but not the keyboard', () => {
		present();
		move(1);
		expect(latest!.presentationSlideIndex).toBe(2);
		move(1, 'click');
		expect(latest!.presentationSlideIndex).toBe(2);
		move(1, 'explicit');
		expect(latest!.presentationSlideIndex).toBe(3);
	});

	it('stays on the first show slide on a back press', () => {
		present();
		move(-1);
		expect(latest!.presentationSlideIndex).toBe(0);
	});
});
