import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createFindReplaceController } from './find-replace-controller';

const text = (id: string, value: string): PptxElement =>
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

function setup() {
	let slides: readonly PptxSlide[] = [
		{ id: 'a', rId: 'a', slideNumber: 1, elements: [text('a1', 'foo bar foo')] },
	];
	const navigate = vi.fn();
	const controller = createFindReplaceController({
		slides: () => slides,
		commit: (next) => {
			slides = next;
		},
		navigate,
	});
	return { controller, navigate, slides: () => slides };
}

describe('find and replace controller', () => {
	it('toggles with Ctrl+F, switches mode with Ctrl+H, and closes on the same chord', () => {
		const { controller } = setup();
		expect(controller.handleShortcut('find')).toBe(true);
		expect(controller.getSnapshot()).toMatchObject({ open: true, mode: 'find' });
		controller.handleShortcut('findReplace');
		expect(controller.getSnapshot()).toMatchObject({ open: true, mode: 'replace' });
		controller.handleShortcut('findReplace');
		expect(controller.getSnapshot().open).toBe(false);
		expect(controller.handleShortcut('copy')).toBe(false);
	});

	it('drops the matches on close and searches the kept query again on open', () => {
		const { controller, navigate } = setup();
		controller.open();
		controller.search({ query: 'foo' });
		expect(controller.getSnapshot().session.matches).toHaveLength(2);
		controller.close();
		expect(controller.getSnapshot().session).toMatchObject({ query: 'foo', matches: [] });
		navigate.mockClear();
		controller.toggle();
		expect(controller.getSnapshot().session.matches).toHaveLength(2);
		expect(navigate).toHaveBeenCalled();
	});

	it('replaces through the shared session and follows a two-way open flag', () => {
		const { controller, slides } = setup();
		controller.setOpen(true);
		controller.search({ query: 'foo' });
		controller.setReplacement('baz');
		expect(controller.replaceAll()).toBe(2);
		expect((slides()[0]!.elements[0] as { text: string }).text).toBe('baz bar baz');
		controller.setOpen(false);
		expect(controller.getSnapshot().open).toBe(false);
	});
});
