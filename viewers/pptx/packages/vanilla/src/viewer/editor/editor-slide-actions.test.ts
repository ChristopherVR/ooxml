import type {
	ChartPptxElement,
	PptxElement,
	PptxHandler,
	PptxSlide,
	PptxThemeColorScheme,
} from 'pptx-viewer-core';
import { createChartElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createEditorOps } from './editor-operations';
import { createSlideActions } from './editor-slide-actions';

const buildSlides = (count: number): PptxSlide[] =>
	Array.from({ length: count }, (_, index) => ({
		id: `slide-${index + 1}`,
		rId: `rId-${index + 1}`,
		slideNumber: index + 1,
		elements: [],
	}));

describe('duplicateSlide', () => {
	it('retains package source identity and isolates chart data before saving', () => {
		const original = buildSlides(1)[0];
		original.id = 'ppt/slides/slide1.xml';
		original.elements = [
			createChartElement('line', { categories: ['A'], series: [{ name: 'One', values: [1] }] }),
		];
		const store = createStore({
			...createInitialViewerState(),
			slides: [original],
			editable: true,
		});
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		createSlideActions({ store, ops, getHandler: () => null }).duplicateSlide();
		const copy = store.get().slides[1];
		expect(copy.sourceSlideId).toBe(original.id);
		(copy.elements[0] as ChartPptxElement).chartData!.series[0].values[0] = 2;
		expect((original.elements[0] as ChartPptxElement).chartData!.series[0].values[0]).toBe(1);
	});
});

describe('addSlide then undo', () => {
	it('brings the stage back to a slide that exists when the new slide was the last one', () => {
		const store = createStore({
			...createInitialViewerState(),
			slides: buildSlides(3),
			currentSlide: 2,
			editable: true,
		});
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		createSlideActions({ store, ops, getHandler: () => null }).addSlide();
		expect(store.get().slides).toHaveLength(4);
		expect(store.get().currentSlide).toBe(3);

		ops.undo();

		// Before the fix `currentSlide` stayed 3 on a three-slide deck, so the
		// stage looked up a slide that no longer existed and drew nothing.
		expect(store.get().slides).toHaveLength(3);
		expect(store.get().currentSlide).toBe(2);
		ops.redo();
		expect(store.get().slides).toHaveLength(4);
		expect(store.get().currentSlide).toBe(2);
	});
});

describe('insertSlideFromTemplate', () => {
	it('inserts the built template after the current slide, selected and history-integrated', () => {
		const store = createStore({
			...createInitialViewerState(),
			slides: buildSlides(2),
			currentSlide: 0,
			editable: true,
		});
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		const pushHistory = vi.spyOn(ops, 'pushHistory');
		const commitChange = vi.spyOn(ops, 'commitChange');
		const actions = createSlideActions({ store, ops, getHandler: () => null });

		actions.insertSlideFromTemplate('agenda');

		const state = store.get();
		expect(state.slides).toHaveLength(3);
		expect(state.currentSlide).toBe(1);
		// A non-blank template inserts real starter content.
		expect(state.slides[1].elements.length).toBeGreaterThan(0);
		// The deck stays renumbered to match array order.
		expect(state.slides.map((slide) => slide.slideNumber)).toStrictEqual([1, 2, 3]);
		expect(state.selectedElementId).toBeNull();
		expect(state.selectedElementIds).toStrictEqual([]);
		expect(pushHistory).toHaveBeenCalledOnce();
		expect(commitChange).toHaveBeenCalledOnce();
	});

	it('pushes history BEFORE the mutation so undo restores the previous deck', () => {
		const store = createStore({
			...createInitialViewerState(),
			slides: buildSlides(2),
			currentSlide: 0,
			editable: true,
		});
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		const actions = createSlideActions({ store, ops, getHandler: () => null });

		actions.insertSlideFromTemplate('title');
		expect(store.get().slides).toHaveLength(3);
		ops.undo();
		expect(store.get().slides).toHaveLength(2);
		expect(store.get().slides.map((slide) => slide.id)).toStrictEqual(['slide-1', 'slide-2']);
	});

	it('is a no-op when editing is disabled', () => {
		const store = createStore({ ...createInitialViewerState(), slides: buildSlides(2) });
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		const pushHistory = vi.spyOn(ops, 'pushHistory');
		const actions = createSlideActions({ store, ops, getHandler: () => null });

		actions.insertSlideFromTemplate('quote');

		expect(store.get().slides).toHaveLength(2);
		expect(pushHistory).not.toHaveBeenCalled();
	});
});

describe('insertSlideFromLayout', () => {
	it('renders the fetched inherited artwork on the inserted slide', async () => {
		const store = createStore({
			...createInitialViewerState(),
			slides: buildSlides(2),
			currentSlide: 0,
			editable: true,
		});
		const artwork = [{ id: 'layout-art', type: 'shape' } as PptxElement];
		const handler = {
			applyLayoutToSlide: vi.fn(
				async (index: number, _path: string, slides: PptxSlide[]) =>
					({
						...slides[index],
						elements: [{ id: 'placeholder', type: 'shape' } as PptxElement],
					}) as PptxSlide,
			),
			getTemplateElementsForSlide: vi.fn().mockResolvedValue(artwork),
		} as unknown as PptxHandler;
		const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: vi.fn() });
		const actions = createSlideActions({ store, ops, getHandler: () => handler });

		actions.insertSlideFromLayout('ppt/slideLayouts/slideLayout1.xml');
		await vi.waitFor(() =>
			expect(Object.keys(store.get().templateElementsBySlideId)).toHaveLength(1),
		);

		const inserted = store.get().slides[1]!;
		expect(handler.getTemplateElementsForSlide).toHaveBeenCalledWith(inserted.id);
		expect(inserted.elements.map((element) => element.id)).toStrictEqual(['placeholder']);
		expect(store.get().templateElementsBySlideId[inserted.id]).toStrictEqual(artwork);
	});
});

describe('undo of Home > Layout and of New Slide from a layout', () => {
	const L1 = 'ppt/slideLayouts/slideLayout1.xml';
	const L2 = 'ppt/slideLayouts/slideLayout2.xml';
	const L3 = 'ppt/slideLayouts/slideLayout3.xml';

	/** The core contract: the slide at `index` is replaced in the array it is given, and returned. */
	function inPlaceDeck() {
		const handler = {
			applyLayoutToSlide: vi.fn(async (index: number, layoutPath: string, slides: PptxSlide[]) => {
				const updated = { ...slides[index]!, layoutPath, layoutName: 'Remapped' };
				slides[index] = updated;
				return updated;
			}),
			getTemplateElementsForSlide: vi.fn().mockResolvedValue([]),
		} as unknown as PptxHandler;
		const store = createStore({
			...createInitialViewerState(),
			slides: buildSlides(2).map((slide) => ({ ...slide, layoutPath: L1 })),
			currentSlide: 0,
			editable: true,
		});
		const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: vi.fn() });
		const actions = createSlideActions({ store, ops, getHandler: () => handler });
		return { store, ops, actions };
	}

	it('one Undo takes back one layout pick', async () => {
		const { store, ops, actions } = inPlaceDeck();
		const layout = () => store.get().slides[0]?.layoutPath;
		actions.applyLayout(L2);
		await vi.waitFor(() => expect(layout()).toBe(L2));
		actions.applyLayout(L3);
		await vi.waitFor(() => expect(layout()).toBe(L3));
		ops.undo();
		expect(layout()).toBe(L2);
		ops.undo();
		expect(layout()).toBe(L1);
		expect(ops.canUndo()).toBeFalsy();
	});

	it('one Undo takes back a new slide from a layout', async () => {
		const { store, ops, actions } = inPlaceDeck();
		actions.insertSlideFromLayout(L2);
		await vi.waitFor(() => expect(store.get().slides[1]?.layoutName).toBe('Remapped'));
		expect(store.get().slides).toHaveLength(3);
		ops.undo();
		expect(store.get().slides.map((slide) => slide.id)).toStrictEqual(['slide-1', 'slide-2']);
		expect(ops.canUndo()).toBeFalsy();
	});
});

describe('getTemplateScheme', () => {
	it('maps the loaded theme colour scheme into a template scheme', () => {
		const colorScheme: PptxThemeColorScheme = {
			dk1: '#111111',
			lt1: '#FEFEFE',
			dk2: '#222222',
			lt2: '#EEEEEE',
			accent1: '#AA0000',
			accent2: '#00AA00',
			accent3: '#0000AA',
			accent4: '#AAAA00',
			accent5: '#00AAAA',
			accent6: '#AA00AA',
			hlink: '#0563C1',
			folHlink: '#954F72',
		};
		const store = createStore({ ...createInitialViewerState(), colorScheme });
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		const actions = createSlideActions({ store, ops, getHandler: () => null });

		expect(actions.getTemplateScheme()).toMatchObject({ accent1: '#AA0000', dk1: '#111111' });
	});

	it('returns an empty map before a theme is loaded (builders fall back to defaults)', () => {
		const store = createStore(createInitialViewerState());
		const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
		const actions = createSlideActions({ store, ops, getHandler: () => null });

		expect(actions.getTemplateScheme()).toStrictEqual({});
	});
});
