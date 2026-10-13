import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { expect, test } from 'vitest';

import { applyElementAnimationStyles } from './animation/animation-dom';
import { createTranslator } from './i18n';
import { createDefaultRegistry, renderSlideStage } from './render';

test('updates background paint independently without hiding the shape text', () => {
	const element = {
		type: 'shape',
		id: 's',
		x: 10,
		y: 20,
		width: 200,
		height: 100,
		shapeType: 'rect',
		shapeStyle: { fillColor: '#ff0000' },
		text: 'Persistent text',
		textSegments: [{ text: 'Persistent text', style: {} }],
	} as PptxElement;
	const states = new Map([['s::pptx-bg', { visible: false, cssAnimation: 'fade 1s both' }]]);
	const stage = renderSlideStage({
		document,
		slide: { id: 'slide', rId: 'rId1', slideNumber: 1, elements: [element] } as PptxSlide,
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator('en'),
		scale: 1,
		interactive: false,
		presenting: true,
		presentationStates: states,
	});
	const root = stage.querySelector<HTMLElement>('[data-element-id="s"]')!;
	const paint = root.querySelector<HTMLElement>('[data-pptx-animation-layer="background"]')!;
	expect(paint.style.visibility).toBe('hidden');
	expect(root.style.backgroundColor).toBe('');
	expect(root.textContent).toContain('Persistent text');
	expect(paint.textContent).not.toContain('Persistent text');
	applyElementAnimationStyles(
		stage,
		new Map([['s::pptx-bg', { visible: true, cssAnimation: 'fade 1s both' }]]),
		new Set(),
		new Set(),
	);
	expect(paint.style.visibility).not.toBe('hidden');
	expect(paint.style.animation).toContain('fade');
	expect(root.style.visibility).not.toBe('hidden');
});
