import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import SlideStage from './SlideStage.svelte';

/**
 * A group child's strokes snap to the device grid by their ABSOLUTE position
 * (group origin plus the child's offset), as PowerPoint's export does. A
 * connector never reads its `parentGroupFill` prop, so the group renderer has to
 * evaluate the shared `getGroupChildParentFill` (which links the children to the
 * group) before the child exists.
 */
let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

describe('svelte: group child strokes snap by absolute position', () => {
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
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(SlideStage, {
			target,
			props: {
				slide: { id: 's1', rId: 'r1', slideNumber: 1, elements: [group] } as PptxSlide,
				canvasSize: { width: 960, height: 540 },
				mediaDataUrls: new Map<string, string>(),
				scale: 0.5,
				interactive: true,
			},
		});
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
		};
		const style = [
			...target.querySelectorAll('[data-element-id="c1"] path, [data-element-id="c1"] line'),
		]
			.map((node) => node.getAttribute('style') ?? '')
			.join('\n');
		expect(style).toContain('round(123.4px');
		expect(style).not.toContain('round(23.4px');
	});
});
