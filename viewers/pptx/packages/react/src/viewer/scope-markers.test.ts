/**
 * The stylesheet's Tailwind preflight matches only inside `[data-pptx-viewer]`,
 * `[data-pptx-viewport]` and `[data-pptx-scope]`, and an outermost marked
 * element takes the page's base text. On a shared shadow-DOM control
 * (`pptx-ui-context-menu`, `pptx-ui-options-dialog`, ...) that text would beat
 * the control's own `:host` typography, so a marker always goes on an ordinary
 * element around the control.
 */
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import {
	findScopedCustomElements,
	scopedCustomElementsIn,
} from '../../../../scripts/scope-marker-tags.mjs';

describe('react scope markers', () => {
	it('never sit on a custom element', () => {
		const src = path.resolve(import.meta.dirname, '..');
		expect(scopedCustomElementsIn(src, ['.tsx', '.ts'])).toStrictEqual([]);
	});

	it('are found on a custom element in JSX', () => {
		expect(
			findScopedCustomElements("<pptx-ui-context-menu ref={ref} data-pptx-scope='' />"),
		).toHaveLength(1);
		expect(
			findScopedCustomElements('<office-ui-dialog data-pptx-viewer="">x</office-ui-dialog>'),
		).toHaveLength(1);
		expect(
			findScopedCustomElements(
				"<div data-pptx-scope='' style={{ display: 'contents' }}><pptx-ui-context-menu ref={ref} /></div>",
			),
		).toStrictEqual([]);
		// A selector in an attribute value is not a marker.
		expect(
			findScopedCustomElements("<pptx-ui-x data-target={el.closest('[data-pptx-viewport]')} />"),
		).toStrictEqual([]);
	});
});
