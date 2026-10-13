import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { expect, test } from 'vitest';

import { applyElementAnimationStyles } from './animation/animation-dom';
import { createTranslator } from './i18n';
import { createDefaultRegistry, renderSlideStage } from './render';

test('applies a fixed whole opacity and restores authored alpha with no CSS animation', () => {
	const root = document.createElement('div');
	const shape = document.createElement('div');
	shape.dataset.elementId = 's';
	shape.style.opacity = '0.6';
	root.append(shape);
	applyElementAnimationStyles(
		root,
		new Map([['s', { visible: true, cssAnimation: undefined, opacity: 0.4 }]]),
		new Set(),
		new Set(),
	);
	expect(shape.style.opacity).toBe('0.4');
	expect(shape.style.animation).toBe('');
	applyElementAnimationStyles(root, new Map(), new Set(), new Set());
	expect(shape.style.opacity).toBe('0.6');
});

test('updates background paint independently without hiding the shape text', () => {
	const element = {
		type: 'shape',
		id: 's',
		x: 10,
		y: 20,
		width: 200,
		height: 100,
		shapeType: 'rect',
		shapeStyle: { fillColor: '#ff0000', strokeColor: '#0000ff', strokeWidth: 4 },
		text: 'Persistent text',
		textSegments: [{ text: 'Persistent text', style: {} }],
	} as PptxElement;
	const states = new Map([
		[
			's::pptx-bg',
			{ visible: false, cssAnimation: 'fade 1s both', animatesFill: true, animatesStroke: true },
		],
	]);
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
	expect(paint.style.backgroundColor).toBe('');
	expect(paint.style.getPropertyValue('--pptx-animation-fill-base')).toBe('#ff0000');
	expect(paint.style.getPropertyValue('--pptx-animation-stroke-base')).toBe('#0000ff');
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
