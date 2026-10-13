import { describe, expect, it } from 'vitest';
import type { PptxSlide } from 'pptx-viewer-core';

import {
	applySorterAction,
	createSlideSorterState,
	selectSorterSlide,
	sorterMenuContext,
	sorterSelectionIndexes,
} from './slide-sorter-state';

const slides = [{ id: 'a' }, { id: 'b', hidden: true }, { id: 'c' }, { id: 'd' }];
describe('sorter state', () => {
	it.each(['ctrlKey', 'metaKey'] as const)(
		'extends a range with %s + Shift without toggling its endpoint',
		(modifier) => {
			let state = selectSorterSlide(createSlideSorterState(slides, 0), slides, 2, {
				[modifier]: true,
			});
			state = selectSorterSlide(state, slides, 0, { [modifier]: true, shiftKey: true });
			expect(sorterSelectionIndexes(state, slides)).toStrictEqual([0, 1, 2]);
			expect(state.anchor).toBe(2);
		},
	);

	it('preserves a separate selection when adding a range', () => {
		let state = selectSorterSlide(createSlideSorterState(slides, 0), slides, 2, { ctrlKey: true });
		state = selectSorterSlide(state, slides, 3, { ctrlKey: true, shiftKey: true });
		expect(sorterSelectionIndexes(state, slides)).toStrictEqual([0, 2, 3]);
		state = selectSorterSlide(state, slides, 1, { shiftKey: true });
		expect(sorterSelectionIndexes(state, slides)).toStrictEqual([1, 2]);
	});
	it('keeps an anchor for ranges, toggles ids, and preserves a right-clicked selection', () => {
		let state = createSlideSorterState(slides, 1);
		state = selectSorterSlide(state, slides, 3, { shiftKey: true });
		expect(sorterSelectionIndexes(state, slides)).toStrictEqual([1, 2, 3]);
		expect(selectSorterSlide(state, slides, 2, {}, true)).toBe(state);
		state = selectSorterSlide(state, slides, 2, { metaKey: true });
		expect(state.selectedIds).toStrictEqual(['b', 'd']);
	});

	it('pastes retained clipboard snapshots after the destination, even after source deletion', () => {
		let state = selectSorterSlide(createSlideSorterState(slides, 0), slides, 2, { ctrlKey: true });
		const source: PptxSlide[] = slides.map((item, index) => ({
			...item,
			rId: item.id,
			slideNumber: index + 1,
			elements: [],
		}));
		state = applySorterAction(state, slides, 'copy', 0, source).state;
		state = selectSorterSlide(state, slides, 1);
		const result = applySorterAction(state, [slides[1]!], 'paste', 0);
		expect(result.paste?.sourceIds).toEqual(['a', 'c']);
		expect(result.paste?.insertAt).toBe(1);
		expect(result.state.selectedIds).toEqual(result.paste!.copies.map((copy) => copy.id));
		expect(sorterMenuContext(state, [slides[1]!]).hasClipboard).toBe(true);
	});

	it('protects the last slide, collapses Escape and clamps zoom', () => {
		const state = applySorterAction(
			createSlideSorterState(slides, 0),
			slides,
			'selectAll',
			0,
		).state;
		expect(sorterMenuContext(state, slides).wouldDeleteAllSlides).toBeTruthy();
		expect(applySorterAction(state, slides, 'delete', 0).indexes).toStrictEqual([]);
		expect(
			applySorterAction(state, slides, 'collapseSelection', 2).state.selectedIds,
		).toStrictEqual(['c']);
		expect(applySorterAction({ ...state, zoom: 200 }, slides, 'zoomIn', 0).state.zoom).toBe(200);
	});
});
