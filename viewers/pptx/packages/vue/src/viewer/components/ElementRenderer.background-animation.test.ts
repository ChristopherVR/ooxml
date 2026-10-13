import { mount } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { expect, test } from 'vitest';
import { shallowRef } from 'vue';

import { PresentationElementStatesKey } from '../composables/presentation-element-states';
import ElementRenderer from './ElementRenderer.vue';

test('animates the paint layer independently from the shape text and resets to static paint', async () => {
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
	const states = shallowRef(
		new Map([['s::pptx-bg', { visible: false, cssAnimation: 'fade 1s both' }]]),
	);
	const wrapper = mount(ElementRenderer, {
		props: { element, mediaDataUrls: new Map(), zIndex: 0 },
		global: { provide: { [PresentationElementStatesKey as symbol]: states } },
	});
	try {
		const root = wrapper.get('[data-element-id="s"]').element as HTMLElement;
		const paint = root.querySelector<HTMLElement>('[data-pptx-animation-layer="background"]')!;
		expect(paint.style.visibility).toBe('hidden');
		expect(paint.style.animation).toContain('fade');
		expect(root.style.backgroundColor).toBe('');
		expect(root.textContent).toContain('Persistent text');
		expect(paint.textContent).not.toContain('Persistent text');
		states.value = new Map();
		await wrapper.vm.$nextTick();
		expect(root.querySelector('[data-pptx-animation-layer="background"]')).toBeNull();
		expect(root.style.backgroundColor).not.toBe('');
	} finally {
		wrapper.unmount();
	}
});
