import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ShapeEffectOverlay from './ShapeEffectOverlay.svelte';

/**
 * A resizing shape's Morph re-paths its outline per frame (`morph-shape-carry`
 * in ooxml-ui), through a descendant rule on `[data-pptx-outline] path`. The
 * outline overlay has to carry that marker in every binding or the outline
 * scales with the box instead.
 */
let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

const card = {
	id: 'ppt/slides/slide2.xml-shape-0',
	type: 'shape',
	x: 0,
	y: 0,
	width: 400,
	height: 160,
	shapeType: 'roundRect',
	shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
} as unknown as PptxElement;

const tri = {
	...card,
	shapeStyle: { ...card.shapeStyle, strokeWidth: 12, compoundLine: 'tri' },
} as unknown as PptxElement;

describe('shapeEffectOverlay morph outline marker', () => {
	it('stamps data-pptx-outline on the outline overlay that holds the path', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		mounted = mount(ShapeEffectOverlay, {
			target,
			props: { element: card, mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		expect(target.querySelector('svg[data-pptx-outline] path')).not.toBeNull();
	});

	it('paints each strand of a compound outline as a sibling path, in order', () => {
		// The compound carry addresses strands as `[data-pptx-outline] path:nth-of-type(n)`.
		const target = document.createElement('div');
		document.body.appendChild(target);
		mounted = mount(ShapeEffectOverlay, {
			target,
			props: { element: tri, mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		expect(target.querySelectorAll('svg[data-pptx-outline] > path')).toHaveLength(3);
	});
});
