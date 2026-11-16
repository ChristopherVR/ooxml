// @vitest-environment jsdom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { buildOutline, OUTLINE_ROW_ATTR } from './outline-view';
import { createOutlineViewSession, focusOutlineRow } from './outline-view-session';

function title(id: string, text: string): PptxElement {
	return {
		type: 'text',
		id,
		name: 'Title',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		text,
		rawXml: { 'p:nvSpPr': { 'p:nvPr': { 'p:ph': { '@_type': 'title' } } } },
	} as PptxElement;
}

function setup(canEdit = true) {
	let slides: PptxSlide[] = [
		{ id: 's1', rId: '', slideNumber: 1, elements: [title('t1', 'Agenda')] },
		{ id: 's2', rId: '', slideNumber: 2, elements: [title('t2', 'Close')] },
	];
	const commit = vi.fn((next: PptxSlide[]) => {
		slides = next;
	});
	const session = createOutlineViewSession({
		slides: () => slides,
		canvasSize: () => ({ width: 960, height: 540 }),
		canEdit: () => canEdit,
		commit,
	});
	return { session, commit, rows: () => buildOutline(slides) };
}

function drawRows(keys: readonly string[]): HTMLElement {
	const container = document.createElement('div');
	for (const key of keys) {
		const input = document.createElement('input');
		input.value = key;
		input.setAttribute(OUTLINE_ROW_ATTR, key);
		container.append(input);
	}
	document.body.append(container);
	return container;
}

afterEach(() => document.body.replaceChildren());

describe('outline view session', () => {
	it('commits a text edit and leaves a no-op alone', () => {
		const { session, commit, rows } = setup();
		const key = rows()[0]!.key;
		expect(session.setText(key, 'Agenda today')).toBe(true);
		expect(commit).toHaveBeenCalledTimes(1);
		expect(rows()[0]!.text).toBe('Agenda today');
		expect(session.setText(key, 'Agenda today')).toBe(false);
		expect(session.setText('missing', 'x')).toBe(false);
		expect(commit).toHaveBeenCalledTimes(1);
	});

	it('refuses every edit on a read-only deck', () => {
		const { session, commit, rows } = setup(false);
		expect(session.setText(rows()[0]!.key, 'Changed')).toBe(false);
		expect(commit).not.toHaveBeenCalled();
	});

	it('maps Enter on a title to a new slide, prevents the default and hands the caret on', () => {
		const { session, commit, rows } = setup();
		const event = { key: 'Enter', preventDefault: vi.fn() };
		session.handleKey(event, rows()[0]!.key);
		expect(event.preventDefault).toHaveBeenCalled();
		expect(commit).toHaveBeenCalledTimes(1);
		expect(commit.mock.calls[0]![0]).toHaveLength(3);
		const pending = session.pendingFocusKey();
		expect(pending).not.toBeNull();
		const container = drawRows(rows().map((row) => row.key));
		session.restoreFocus(container);
		expect(document.activeElement?.getAttribute(OUTLINE_ROW_ATTR)).toBe(pending);
		expect(session.pendingFocusKey()).toBeNull();
	});

	it('leaves keys the outline does not map alone', () => {
		const { session, commit, rows } = setup();
		const event = { key: 'a', preventDefault: vi.fn() };
		session.handleKey(event, rows()[0]!.key);
		expect(event.preventDefault).not.toHaveBeenCalled();
		expect(commit).not.toHaveBeenCalled();
	});

	it('finds a row whose key carries selector characters, caret at the end', () => {
		const key = 's1|t1 "0"';
		const container = drawRows(['other', key]);
		expect(focusOutlineRow(container, key)).toBe(true);
		const input = document.activeElement as HTMLInputElement;
		expect(input.getAttribute(OUTLINE_ROW_ATTR)).toBe(key);
		expect(input.selectionStart).toBe(key.length);
		expect(focusOutlineRow(container, 'gone')).toBe(false);
	});

	it('keeps nothing pending once restored, even without a container', () => {
		const { session, rows } = setup();
		session.setText(rows()[0]!.key, 'Again');
		session.restoreFocus(null);
		expect(session.pendingFocusKey()).not.toBeNull();
		session.restoreFocus(drawRows([]));
		expect(session.pendingFocusKey()).toBeNull();
	});
});
