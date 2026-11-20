// @vitest-fresh-modules: the lazy layout library must not have loaded when the first test runs,
// so this file gets a module registry of its own instead of resetting a shared one.
import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildRibbonGallery } from './gallery-registry';

const smartArt = {
	id: 'sa1',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	smartArtData: {
		layout: 'basicBlockList',
		resolvedLayoutType: 'list',
		nodes: [{ id: 'n1', text: 'Node One' }],
		connections: [],
	},
} as unknown as PptxElement;

describe('Layouts gallery before the lazy layout library has loaded', () => {
	it('lists only the families and hands out a promise for the named layouts', async () => {
		const first = buildRibbonGallery('smartArtLayouts', { element: smartArt });
		expect(first.sections.map((section) => section.id)).toEqual(['layouts']);
		expect(first.ready).toBeInstanceOf(Promise);

		const next = await first.ready;
		expect(next?.sections.map((section) => section.id)).toEqual(
			expect.arrayContaining(['layouts', 'named-timeline']),
		);
		// Built once the library is in, so there is nothing left to wait for.
		expect(next?.ready).toBeUndefined();
	});

	it('offers no promise once the library is available', async () => {
		await buildRibbonGallery('smartArtLayouts', { element: smartArt }).ready;
		expect(buildRibbonGallery('smartArtLayouts', { element: smartArt }).ready).toBeUndefined();
	});
});
