// @vitest-environment happy-dom
/**
 * A Design gallery theme pick made while that tile's hover preview shows.
 *
 * The gallery sends the pick while the preview still stands, so the host first puts the original
 * slides back and then commits. React's theme handlers took the theme to re-resolve from out of
 * the last render, which during a preview is the PREVIEWED theme: the commit then mapped the
 * previewed colours and faces onto themselves, the slides kept the original look, and the theme
 * said the pick had applied.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { applyRibbonGalleryItem } from 'ooxml-ui/pptx';
import type { RibbonGalleryId } from 'ooxml-ui/pptx';
import { act } from 'react';
import { beforeAll, describe, expect, it } from 'vitest';

import {
	betaTitleFont,
	flush,
	flushUntil,
	live,
	withHarness,
} from './theme-apply-harness.test-support';

let deckBytes: Uint8Array;

beforeAll(() => {
	deckBytes = new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/field-substitution.pptx')),
	);
});

/** Hover `itemId` until its preview shows, then pick it, as a pointer does. */
async function hoverThenPick(
	gallery: RibbonGalleryId,
	itemId: string,
	face: string,
): Promise<void> {
	await flushUntil(() => live().parses >= 1 && !live().loading);
	await act(async () => {
		live().setActiveSlideIndex(1);
	});
	await flush();
	const result = applyRibbonGalleryItem(gallery, itemId, live().gallery.context);
	expect(result).not.toBeNull();
	await act(async () => {
		live().gallery.preview?.(result);
	});
	await flush();
	expect(betaTitleFont(), 'the preview shows on the slide').toBe(face);
	expect(live().history.canUndo, 'a preview is not an undo step').toBeFalsy();
	await act(async () => {
		live().gallery.dispatch(result!);
	});
	// Bounded well inside the test timeout, so a pick that records no undo step fails below.
	await flushUntil(() => live().history.canUndo, 2000);
	await flush(100);
}

describe('a gallery theme pick made during its hover preview', () => {
	it.each([
		['themes', 'facet', 'Trebuchet MS'],
		['themeFonts', 'arial', 'Arial'],
	] as const)(
		'%s > %s applies to the slides it was previewed on',
		async (gallery, itemId, face) => {
			await withHarness(deckBytes, async () => {
				await hoverThenPick(gallery, itemId, face);
				expect({
					title: betaTitleFont(),
					theme: live().theme?.fontScheme?.majorFont?.latin,
					undoable: live().history.canUndo,
				}).toStrictEqual({
					title: face,
					theme: face,
					undoable: true,
				});
				await act(async () => {
					live().history.handleUndo();
				});
				await flush(50);
				expect(betaTitleFont()).toBe('Calibri Light');
			});
		},
	);
});
