import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { getShapeFillStrokeStyle } from '../style/element-style';
import ShapeEffectOverlay from './ShapeEffectOverlay.svelte';

/**
 * A non-rectangular preset paints its fill as a clip-path, which also clipped
 * its shadow and glow. It now paints them as one shape-following SVG filter.
 */
let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

function triangle(): PptxElement {
	return {
		id: 'tri-1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 160,
		height: 130,
		shapeType: 'triangle',
		shapeStyle: {
			fillColor: '#2070C0',
			fillMode: 'solid',
			shadowColor: '#000000',
			shadowOpacity: 0.5,
			shadowBlur: 10,
			shadowAngle: 45,
			shadowDistance: 8,
			outerShadowXml: {},
		},
	} as unknown as PptxElement;
}

describe('shape-following shadow (svelte)', () => {
	it('does not clip the container and puts the shadow in a filter', () => {
		const style = getShapeFillStrokeStyle(triangle());
		expect(style.clipPath).toBeUndefined();
		expect(style.boxShadow).toBeUndefined();
		expect(style.filter).toContain('url(#shape-fx-tri-1)');
		expect(style.backgroundColor).toBe('transparent');
	});

	it('injects the filter and paints the fill as an SVG path', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		mounted = mount(ShapeEffectOverlay, {
			target,
			props: { element: triangle(), mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		const html = target.innerHTML;
		expect(html).toContain('id="shape-fx-tri-1"');
		expect(html).toContain('pptx-svelte-subpath-fill');
		expect(html).toContain('fill="#2070C0"');
	});
});
