import { Font, Glyph, Path } from 'opentype.js';
import { expect, test, vi } from 'vitest';

import { glyphOutlineFontCache } from './glyph-outline-cache';
import { PptxViewer } from './PptxViewer';

function fontBytes(): Uint8Array {
	const path = new Path();
	path.moveTo(100, 0);
	path.lineTo(100, 700);
	path.lineTo(300, 700);
	path.close();
	const font = new Font({
		familyName: 'VanillaLazyOutline',
		styleName: 'Regular',
		unitsPerEm: 1000,
		ascender: 800,
		descender: -200,
		glyphs: [
			new Glyph({ name: '.notdef', unicode: 0, advanceWidth: 650, path: new Path() }),
			new Glyph({ name: 'A', unicode: 65, advanceWidth: 650, path }),
		],
	});
	return new Uint8Array(font.toArrayBuffer());
}

test('repaints mounted envelope glyphs when outlines become ready, until destroyed', async () => {
	const container = document.createElement('div');
	document.body.append(container);
	const viewer = new PptxViewer(container);
	try {
		viewer.store.set({
			slides: [
				{
					id: 's1',
					rId: 'r1',
					slideNumber: 1,
					elements: [
						{
							id: 'outline',
							type: 'text',
							x: 10,
							y: 10,
							width: 200,
							height: 100,
							text: 'A',
							textStyle: { textWarpPreset: 'textInflate', fontFamily: 'VanillaLazyOutline' },
						},
					],
				},
			],
		});
		const outline = () => container.querySelector('[data-element-id="outline"]')?.outerHTML ?? '';
		await vi.waitFor(() => expect(outline()).toContain('<svg'));
		expect(outline()).not.toContain('<path');
		const tick = viewer.store.get().outlineFontsTick;
		// The font arrives after the glyph was painted (a deferred parser load or
		// a webfont fetch): the cache notification alone must repaint it.
		glyphOutlineFontCache.registerFontBytes('VanillaLazyOutline', false, false, fontBytes());
		await glyphOutlineFontCache.preload();
		await vi.waitFor(() => expect(viewer.store.get().outlineFontsTick).toBeGreaterThan(tick));
		await vi.waitFor(() => expect(outline()).toContain('<path'));
		viewer.destroy();
		const finalTick = viewer.store.get().outlineFontsTick;
		glyphOutlineFontCache.get('VanillaLazyOutlineLate', false, false);
		glyphOutlineFontCache.registerFontBytes('VanillaLazyOutlineLate', false, false, fontBytes());
		await new Promise<void>((resolve) => {
			setTimeout(resolve, 0);
		});
		expect(viewer.store.get().outlineFontsTick).toBe(finalTick);
	} finally {
		viewer.destroy();
		container.remove();
	}
});
