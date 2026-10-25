import { createEditorHistoryController } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createFindReplaceActions } from './editor-find-replace-actions';
import type { EditorOps } from './editor-operations';

function slideWithText(text: string): PptxSlide {
	return {
		id: 's1',
		rId: '',
		slideNumber: 1,
		elements: [
			{
				type: 'text',
				id: 'el1',
				x: 0,
				y: 0,
				width: 100,
				height: 50,
				text,
				textSegments: [{ text, style: {} }],
			},
		],
	} as PptxSlide;
}

/** Three matches of "foo" on two slides. */
function twoSlides(): PptxSlide[] {
	const first = slideWithText('foo bar foo');
	const second = slideWithText('a foo');
	return [
		{ ...first, id: 's1', elements: [{ ...first.elements[0]!, id: 'a1' }] },
		{ ...second, id: 's2', elements: [{ ...second.elements[0]!, id: 'b1' }] },
	] as PptxSlide[];
}

const texts = (slides: readonly PptxSlide[]) =>
	slides.map((s) => (s.elements[0] as { text: string }).text);

function fakeOps(): EditorOps {
	return {
		history: createEditorHistoryController(),
		selectedElement: () => undefined,
		select: vi.fn(),
		pushHistory: vi.fn(),
		commitChange: vi.fn(),
		patchGeometry: vi.fn(),
		previewElement: vi.fn(),
		deleteSelected: vi.fn(),
		duplicateSelected: () => null,
		nudgeSelected: vi.fn(),
		commitInlineText: vi.fn(),
		commitNotes: vi.fn(),
		setHandoutSlidesPerPage: vi.fn(),
		applyFormatPainter: vi.fn(() => false),
		commitTableCell: vi.fn(),
		updateEquation: vi.fn(),
		updateDocumentProperties: vi.fn(),
		updatePresentationProperties: vi.fn(),
		updateSections: vi.fn(),
		updateHeaderFooter: vi.fn(),
		updateCustomShows: vi.fn(),
		undo: vi.fn(),
		redo: vi.fn(),
		canUndo: () => false,
		canRedo: () => false,
		undoLabel: () => undefined,
		redoLabel: () => undefined,
		clearHistory: vi.fn(),
		setHistoryDepth: vi.fn(),
		save: vi.fn(),
	};
}

describe('editor-find-replace-actions', () => {
	it('counts matches without mutating', () => {
		const store = createStore({
			...createInitialViewerState(),
			slides: [slideWithText('hello world')],
		});
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		expect(actions.search('world', false)).toBe(1);
		expect(store.get().slides[0].elements[0]).toMatchObject({ text: 'hello world' });
	});

	it('replaces the first match and pushes history', () => {
		const store = createStore({
			...createInitialViewerState(),
			editable: true,
			slides: [slideWithText('hello world')],
		});
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		const count = actions.replaceCurrent('world', 'there', false);
		expect(count).toBe(1);
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(ops.commitChange).toHaveBeenCalledOnce();
		expect((store.get().slides[0].elements[0] as { text: string }).text).toBe('hello there');
	});

	it('replaces every match with replaceAll', () => {
		const store = createStore({
			...createInitialViewerState(),
			editable: true,
			slides: [slideWithText('cat cat cat')],
		});
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		const count = actions.replaceAll('cat', 'dog', false);
		expect(count).toBe(3);
		expect((store.get().slides[0].elements[0] as { text: string }).text).toBe('dog dog dog');
	});

	it('is a no-op when not editable', () => {
		const store = createStore({ ...createInitialViewerState(), slides: [slideWithText('hello')] });
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		expect(actions.replaceAll('hello', 'bye', false)).toBe(0);
		expect(ops.pushHistory).not.toHaveBeenCalled();
	});

	it('a search goes to the first match; Next and Previous wrap and report n of N', () => {
		const store = createStore({ ...createInitialViewerState(), slides: twoSlides() });
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		expect(actions.search('foo', false)).toBe(3);
		expect(ops.select).toHaveBeenLastCalledWith('a1');
		expect(actions.status!()).toStrictEqual({ current: 1, total: 3 });
		actions.next!();
		actions.next!();
		expect(store.get().currentSlide).toBe(1);
		expect(ops.select).toHaveBeenLastCalledWith('b1');
		expect(actions.status!()).toStrictEqual({ current: 3, total: 3 });
		actions.next!();
		expect(actions.status!().current).toBe(1);
		expect(store.get().currentSlide).toBe(0);
		actions.previous!();
		expect(actions.status!()).toStrictEqual({ current: 3, total: 3 });
	});

	it('replace acts on the current match, then moves on to the next one', () => {
		const store = createStore({
			...createInitialViewerState(),
			editable: true,
			slides: twoSlides(),
		});
		const ops = fakeOps();
		const actions = createFindReplaceActions({ store, ops });
		actions.search('foo', false);
		actions.next!();
		expect(actions.replaceCurrent('foo', 'qux', false)).toBe(1);
		expect(texts(store.get().slides)).toStrictEqual(['foo bar qux', 'a foo']);
		expect(ops.pushHistory).toHaveBeenCalledOnce();
		expect(actions.status!()).toStrictEqual({ current: 2, total: 2 });
		expect(ops.select).toHaveBeenLastCalledWith('b1');
	});

	it('closing drops the matches and keeps the query for the next search', () => {
		const store = createStore({ ...createInitialViewerState(), slides: twoSlides() });
		const actions = createFindReplaceActions({ store, ops: fakeOps() });
		actions.search('foo', false);
		actions.next!();
		actions.close!();
		expect(actions.status!()).toStrictEqual({ current: 0, total: 0 });
		expect(actions.search('foo', false)).toBe(3);
		expect(actions.status!().current).toBe(1);
	});
});
