import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createFindBarController } from './find-bar-controller';
import { buildFindReplaceActions } from './find-layout';

const text = (id: string, value: string) =>
	({
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 20,
		text: value,
		textSegments: [{ text: value, style: {} }],
	}) as PptxElement;

describe('createFindBarController', () => {
	it('searches on a query change, navigates, replaces and drops matches on close', () => {
		let slides: readonly PptxSlide[] = [
			{ id: 'a', rId: 'r1', slideNumber: 1, elements: [text('a1', 'foo foo')] },
		];
		const navigate = vi.fn();
		const bar = createFindBarController({
			slides: () => slides,
			commit: (next) => {
				slides = next;
			},
			navigate,
		});
		const changed = vi.fn();
		bar.subscribe(changed);
		bar.setOpen(true);
		bar.port().setQuery('foo');
		const actions = buildFindReplaceActions(bar.port(), () => undefined);
		expect(actions.open).toBe(true);
		expect(actions.matches).toHaveLength(2);
		expect(actions.currentIndex).toBe(0);
		expect(navigate).toHaveBeenCalledWith({ slideIndex: 0, elementId: 'a1' });
		bar.port().setReplacement('bar');
		expect(buildFindReplaceActions(bar.port(), () => undefined).replaceAll()).toEqual({
			ok: true,
		});
		expect(
			(slides[0]?.elements[0] as { textSegments: Array<{ text: string }> }).textSegments[0]?.text,
		).toBe('bar bar');
		bar.toggle();
		expect(bar.getSnapshot().open).toBe(false);
		expect(bar.getSnapshot().session.matches).toHaveLength(0);
		expect(changed).toHaveBeenCalled();
	});
});
