/**
 * The root display inputs `showToolbar`, `showThumbnails` and `initialSlide`.
 *
 * React, Vue, Svelte and vanilla expose the three; Angular did not, so a host
 * could not hide the chrome, hide the thumbnail pane or open on a given slide
 * here. The defaults and clamping come from the shared
 * `resolveViewerRootOptions` / `resolveInitialSlideIndex` (ooxml-ui/pptx);
 * this package has no TestBed, so the wiring is pinned as source text (the
 * pattern `power-point-viewer-api.test.ts` uses) and the shared resolution is
 * checked at runtime.
 */
import { resolveInitialSlideIndex, resolveViewerRootOptions } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';

const source = componentSource(import.meta.dirname, 'power-point-viewer.component.ts');

/** The `@if (...)` condition that opens the block containing `marker`. */
function gateBefore(marker: string): string {
	const at = source.indexOf(marker);
	expect(at).toBeGreaterThan(-1);
	const open = source.lastIndexOf('@if (', at - 1);
	return source.slice(open, source.indexOf(') {', open) + 1).replace(/\s+/gu, ' ');
}

describe('powerPointViewerComponent root display inputs', () => {
	it('declares the three inputs undefaulted so the shared defaults apply', () => {
		expect(source).toContain('readonly initialSlide = input<number | undefined>(undefined);');
		expect(source).toContain('readonly showToolbar = input<boolean | undefined>(undefined);');
		expect(source).toContain('readonly showThumbnails = input<boolean | undefined>(undefined);');
		expect(source).toMatch(
			/rootOptions = computed\(\(\) =>\s*resolveViewerRootOptions\(\{\s*initialSlide: this\.initialSlide\(\),\s*showToolbar: this\.showToolbar\(\),\s*showThumbnails: this\.showThumbnails\(\),/u,
		);
		expect(resolveViewerRootOptions({})).toStrictEqual({
			initialSlide: 0,
			showToolbar: true,
			showThumbnails: true,
		});
	});

	it('gates the banners, title bar, ribbon, mobile toolbar and status bar on showToolbar', () => {
		expect(source).toMatch(
			/toolbarVisible = computed\(\s*\(\) => this\.chromeVisible\(\) && this\.rootOptions\(\)\.showToolbar,/u,
		);
		expect(gateBefore('pptx-ng-protected-view-banner')).toContain('toolbarVisible()');
		expect(gateBefore('<pptx-readonly-banner')).toContain('toolbarVisible()');
		expect(gateBefore("@if (customizationService.panelVisible('titleBar'))")).toContain('!mobile.isMobile() && toolbarVisible()');
		expect(gateBefore('<pptx-mobile-toolbar')).toContain('toolbarVisible()');
		expect(gateBefore('<pptx-status-bar')).toContain('toolbarVisible()');
		expect(gateBefore('<pptx-mobile-slides-sheet')).toContain('rootOptions().showToolbar');
	});

	it('gates both slide thumbnail panes on showThumbnails', () => {
		expect(source).toMatch(
			/thumbnailsVisible = computed\(\s*\(\) => this\.chromeVisible\(\) && this\.rootOptions\(\)\.showThumbnails,/u,
		);
		expect(gateBefore('<pptx-slides-panel')).toContain('thumbnailsVisible()');
		const readOnlyRail = source.slice(
			source.indexOf('} @else if (', source.indexOf('<pptx-slides-panel')),
			source.indexOf('<nav class="pptx-ng-thumbnails"'),
		);
		expect(readOnlyRail).toContain('thumbnailsVisible()');
	});

	it('opens each loaded deck on the clamped initialSlide', () => {
		expect(source).toContain(
			'this.activeSlideIndex.set(resolveInitialSlideIndex(this.initialSlide(), slides.length));',
		);
		expect(resolveInitialSlideIndex(2, 3)).toBe(2);
		expect(resolveInitialSlideIndex(99, 3)).toBe(2);
		expect(resolveInitialSlideIndex(undefined, 3)).toBe(0);
	});
});
