import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import ShapeEffectOverlay from './ShapeEffectOverlay.svelte';

/**
 * Split-layout presets (an unstroked FILL path plus a separate stroke path: `arc`, `can`,
 * brackets ...). PowerPoint strokes only the stroked path, never the radii of an arc's wedge,
 * and fills the wedge when the shape has a fill.
 */
let mounted: ReturnType<typeof mount> | undefined;

afterEach(() => {
	if (mounted) {
		void unmount(mounted);
		mounted = undefined;
	}
	document.body.innerHTML = '';
});

function arc(): PptxElement {
	return {
		type: 'shape',
		id: 'sp1',
		x: 0,
		y: 0,
		width: 200,
		height: 120,
		shapeType: 'arc',
		shapeAdjustments: { adj1: 16200000, adj2: 0 },
		shapeStyle: { fillColor: '#336699', strokeWidth: 4, strokeColor: '#FF0000' },
	} as unknown as PptxElement;
}

describe('split-layout preset stroke (svelte)', () => {
	it('strokes an arc along its curve only, never the wedge radii', () => {
		const target = document.createElement('div');
		document.body.appendChild(target);
		mounted = mount(ShapeEffectOverlay, {
			target,
			props: { element: arc(), mediaDataUrls: new Map<string, string>(), zIndex: 1 },
		});
		flushSync();
		const paths = [...target.querySelectorAll('path')]
			.filter((p) => p.getAttribute('stroke') === '#FF0000')
			.map((p) => p.getAttribute('d') ?? '');
		expect(paths.length).toBeGreaterThan(0);
		for (const d of paths) {
			expect(d).toContain('A ');
			expect(d).not.toContain('L ');
		}
	});
});
