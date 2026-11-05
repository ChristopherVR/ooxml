import { mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import SlideStage from './SlideStage.vue';

/**
 * A group child's strokes snap to the device grid by their ABSOLUTE position
 * (group origin plus the child's offset), as PowerPoint's export does.
 */
describe('vue: group child strokes snap by absolute position', () => {
	it('snaps a connector inside a group at group y + child y', () => {
		const group = {
			type: 'group',
			id: 'g1',
			x: 200,
			y: 100,
			width: 300,
			height: 100,
			children: [
				{
					type: 'connector',
					id: 'c1',
					x: 10,
					y: 23.4,
					width: 50,
					height: 0,
					shapeStyle: { strokeColor: '#000000', strokeWidth: 0.3333 },
				},
			],
		} as unknown as PptxElement;
		const wrapper = mount(SlideStage, {
			props: {
				slide: { id: 's1', elements: [group] } as unknown as PptxSlide,
				canvasSize: { width: 960, height: 540 },
				mediaDataUrls: new Map<string, string>(),
				scale: 0.5,
			},
		});
		const style = [
			...wrapper.element.querySelectorAll('.pptx-vue-connector path, .pptx-vue-connector line'),
		]
			.map((node) => node.getAttribute('style') ?? '')
			.join('\n');
		expect(style).toContain('round(123.4px');
		expect(style).not.toContain('round(23.4px');
	});
});
