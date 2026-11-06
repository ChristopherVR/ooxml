import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ShapeEffectOverlay from './ShapeEffectOverlay.svelte';

/**
 * A `path="rect"` gradient outline is stroked with a `<pattern>` whose single
 * `<image>` is the nested-rectangle field, in `objectBoundingBox` units. The
 * pattern's CONTENT must be in those units too (`patternContentUnits`): left at
 * its `userSpaceOnUse` default the 1x1 image is one user unit big and the paint
 * is invisible.
 */
let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

function rectPathOutline(): PptxElement {
	return {
		type: 'shape',
		id: 'sp-rect-outline',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		shapeType: 'rect',
		shapeStyle: {
			strokeFillMode: 'gradient',
			strokeWidth: 4,
			strokeColor: '#808080',
			strokeGradientType: 'radial',
			strokeGradientPathType: 'rect',
			strokeGradientStops: [
				{ color: '#FF0000', position: 0 },
				{ color: '#0000FF', position: 100 },
			],
		},
	} as unknown as PptxElement;
}

describe('shapeEffectOverlay rect-path gradient outline', () => {
	it('declares the pattern content in bounding-box units', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		mounted = mount(ShapeEffectOverlay, {
			target,
			props: { element: rectPathOutline(), mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		const pattern = target.querySelector('pattern');
		expect(pattern).not.toBeNull();
		expect(pattern!.getAttribute('patternUnits')).toBe('objectBoundingBox');
		expect(pattern!.getAttribute('patternContentUnits')).toBe('objectBoundingBox');
	});
});
