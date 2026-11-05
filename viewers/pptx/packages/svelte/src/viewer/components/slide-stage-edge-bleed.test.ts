import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import SlideStage from './SlideStage.svelte';

/**
 * Issue #141: two abutting filled rectangles must not show a light seam where
 * their shared edge falls inside a device pixel. The lower box grows under the
 * opaque upper one (shared `resolveStageEdgeBleeds`); a selected box keeps its
 * exact geometry.
 */

function rect(id: string, x: number, fillColor: string): PptxElement {
	return {
		type: 'shape',
		id,
		x,
		y: 0,
		width: 100,
		height: 100,
		shapeType: 'rect',
		text: '',
		shapeStyle: { fillMode: 'solid', fillColor, strokeWidth: 0, strokeFillMode: 'none' },
	} as PptxElement;
}

const slide = {
	id: 's1',
	elements: [rect('lower', 100, '#112233'), rect('upper', 200, '#445566')],
} as unknown as PptxSlide;

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	document.body.innerHTML = '';
});

function boxStyle(id: string, selectedElementIds?: readonly string[]): string {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(SlideStage, {
		target,
		props: {
			slide,
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map<string, string>(),
			interactive: true,
			...(selectedElementIds ? { selectedElementIds } : {}),
		},
	});
	flushSync();
	cleanup = () => void unmount(instance);
	return (
		target.querySelector<HTMLElement>(`[data-element-id="${id}"]`)?.getAttribute('style') ?? ''
	);
}

describe('slideStage abutting-edge bleed', () => {
	it('grows the lower box under the abutting opaque box above it', () => {
		const lower = boxStyle('lower');
		expect(lower).toContain('left: 100px');
		expect(lower).toContain('width: 102px');
	});

	it('leaves the upper box exact', () => {
		expect(boxStyle('upper')).toContain('width: 100px');
	});

	it('keeps the exact box while either box is selected', () => {
		expect(boxStyle('lower', ['upper'])).toContain('width: 100px');
	});
});
