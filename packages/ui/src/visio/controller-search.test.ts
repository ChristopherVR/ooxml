import { describe, expect, it, vi } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';

function model(): VisioDocument {
	const document = structuredClone(demoDocument);
	for (const page of document.pages) {
		page.shapes = [page.shapes[0]!];
		page.shapes[0]!.id = 'same-id';
		page.shapes[0]!.text.plainText = 'Needle needle';
	}
	return document;
}
function searchable(): ViewerController {
	const controller = new ViewerController();
	controller.setDocument(model());
	controller.setSearchQuery('needle');
	return controller;
}
describe('shared controller search and navigation', () => {
	it('searches without navigation, then wraps next/previous with page-scoped identities', () => {
		const controller = searchable(),
			events = vi.fn();
		controller.onEvent(events);
		expect(controller.state.search.results).toHaveLength(2);
		expect(controller.state.selectedShape).toBeNull();
		expect(controller.state.search.activeIndex).toBe(-1);
		controller.previousSearchResult();
		expect(controller.state.pageIndex).toBe(1);
		expect(controller.state.selectedShape?.pageId).toBe('2');
		expect(events.mock.calls.map(([name]) => name)).toEqual(['page-change', 'shape-select']);
		controller.nextSearchResult();
		expect(controller.state.pageIndex).toBe(0);
		expect(controller.state.selectedShape?.pageId).toBe('1');
		controller.nextSearchResult();
		controller.nextSearchResult();
		expect(controller.state.search.activeIndex).toBe(0);
	});
	it('keeps queries and viewport across unrelated changes and deactivates manual navigation', () => {
		const controller = searchable();
		controller.nextSearchResult();
		const search = controller.state.search;
		controller.setZoom(2);
		expect(controller.state.search).toBe(search);
		controller.selectShape(null);
		expect(controller.state.search.activeIndex).toBe(-1);
		controller.nextSearchResult();
		controller.setPage(1);
		expect(controller.state.search.activeIndex).toBe(-1);
		expect(controller.state.search.results).toBe(search.results);
		controller.setSearchQuery('unmatched');
		expect(controller.state.pageIndex).toBe(1);
		expect(controller.state.search.results).toHaveLength(0);
	});
	it('ignores invalid indexes/empty results and rejects overlong queries without changing state', () => {
		const controller = searchable(),
			state = controller.state;
		for (const index of [-1, 2, NaN, Infinity, 0.5]) controller.selectSearchResult(index);
		expect(controller.state).toBe(state);
		expect(() => controller.setSearchQuery('a'.repeat(257))).toThrow('256');
		expect(controller.state).toBe(state);
		controller.setSearchQuery('');
		const cleared = controller.state;
		controller.nextSearchResult();
		controller.previousSearchResult();
		expect(controller.state).toBe(cleared);
	});
	it('clears search on accepted replacement, including the same document, but not rejected input', () => {
		const controller = searchable(),
			state = controller.state;
		const invalid = model();
		invalid.pages[0]!.width = -1;
		expect(() => controller.setDocument(invalid)).toThrow();
		expect(controller.state).toBe(state);
		controller.setDocument(state.document);
		expect(controller.state.search.query).toBe('');
		expect(controller.state.search.results).toHaveLength(0);
		controller.setSearchQuery('needle');
		controller.setDocument(null);
		expect(controller.state.search.results).toHaveLength(0);
	});
	it('retains prior search after failed loading and resets it only after a successful new document', async () => {
		const parser = vi
			.fn()
			.mockRejectedValueOnce(new Error('bad'))
			.mockResolvedValueOnce(demoDocument);
		const controller = new ViewerController(parser);
		controller.setDocument(model());
		controller.setSearchQuery('needle');
		const search = controller.state.search;
		await expect(controller.load(new Uint8Array())).rejects.toThrow('bad');
		expect(controller.state.search).toBe(search);
		await controller.load(new Uint8Array());
		expect(controller.state.search.query).toBe('');
		controller.setSearchQuery('needle');
		expect(controller.state.search.results).toHaveLength(0);
	});
	it('never resets a newer document/query when an older source fails', async () => {
		let reject!: (error: Error) => void;
		const source = new Promise<ArrayBuffer>((_, no) => {
			reject = no;
		});
		const controller = searchable();
		const pending = controller.loadSource(() => source);
		controller.setDocument(demoDocument);
		controller.setSearchQuery('framework');
		const search = controller.state.search;
		reject(new Error('old read'));
		await pending;
		expect(controller.state.search).toBe(search);
	});
	it.each(['replace', 'query', 'navigate', 'destroy'] as const)(
		'aborts stale shape selection when page-change callbacks %s',
		(action) => {
			const controller = searchable(),
				selections = vi.fn();
			let handled = false;
			controller.onEvent((name) => {
				if (name !== 'page-change' || handled) return;
				handled = true;
				if (action === 'replace') controller.setDocument(demoDocument);
				if (action === 'query') controller.setSearchQuery('unmatched');
				if (action === 'navigate') controller.previousSearchResult();
				if (action === 'destroy') controller.destroy();
			});
			controller.onEvent((name, value) => {
				if (name === 'shape-select') selections(value);
			});
			expect(() => controller.selectSearchResult(1)).not.toThrow();
			if (action === 'navigate') {
				expect(controller.state.selectedShape?.pageId).toBe('1');
				expect(selections).toHaveBeenCalledTimes(1);
			} else {
				expect(controller.state.selectedShape).toBeNull();
				expect(selections).not.toHaveBeenCalled();
			}
		},
	);
	it('does not emit stale navigation after a subscriber supersedes it', () => {
		const controller = searchable(),
			events = vi.fn();
		controller.onEvent(events);
		controller.subscribe((state) => {
			if (state.search.activeIndex === 1) controller.setSearchQuery('changed');
		});
		controller.selectSearchResult(1);
		expect(controller.state.search.query).toBe('changed');
		expect(controller.state.selectedShape).toBeNull();
		expect(events).not.toHaveBeenCalled();
	});
	it('isolates viewers and releases all search state on destroy', () => {
		const first = searchable(),
			second = searchable();
		first.nextSearchResult();
		first.destroy();
		expect(first.state.search.results).toHaveLength(0);
		expect(first.state.search.query).toBe('');
		expect(second.state.search.results).toHaveLength(2);
		expect(second.state.search.activeIndex).toBe(-1);
		for (const command of [
			() => first.setSearchQuery(''),
			() => first.selectSearchResult(0),
			() => first.nextSearchResult(),
			() => first.previousSearchResult(),
		])
			expect(command).toThrow('destroyed');
	});
});
