import type { PptxElement } from 'pptx-viewer-core';
import { mount, unmount, flushSync } from 'svelte';
import { expect, test } from 'vitest';

import { PresentationElementStatesKey } from '../state/presentation-element-states-context';
import ElementRenderer from './ElementRenderer.svelte';

test('keeps shape text outside its hidden animated background paint', () => {
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
	const target = document.createElement('div');
	document.body.appendChild(target);
	const states = new Map([
		[
			's::pptx-bg',
			{ visible: false, cssAnimation: 'fade 1s both', animatesFill: true, animatesStroke: true },
		],
	]);
	const instance = mount(ElementRenderer, {
		target,
		props: { element, mediaDataUrls: new Map(), zIndex: 0 },
		context: new Map([[PresentationElementStatesKey, () => states]]),
	});
	try {
		flushSync();
		const root = target.querySelector<HTMLElement>('[data-element-id="s"]')!;
		const paint = root.querySelector<HTMLElement>('[data-pptx-animation-layer="background"]')!;
		expect(paint.style.visibility).toBe('hidden');
		expect(paint.style.backgroundColor).toBe('');
		expect(paint.style.getPropertyValue('--pptx-animation-fill-base')).toBe('#ff0000');
		expect(paint.style.getPropertyValue('--pptx-animation-stroke-base')).toBe('#0000ff');
		expect(paint.style.animation).toContain('fade');
		expect(root.style.backgroundColor).toBe('');
		expect(root.textContent).toContain('Persistent text');
		expect(paint.textContent).not.toContain('Persistent text');
	} finally {
		unmount(instance);
		target.remove();
	}
});
