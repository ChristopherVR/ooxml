// oxlint-disable react-hooks/rules-of-hooks
import { describe, expect, it } from 'vitest';
import { effectScope, nextTick, ref } from 'vue';

import { useSlideChangeSelection } from './useSlideChangeSelection';

function setup(slideCount = 3) {
	const activeSlideIndex = ref(0),
		selectedElementIds = ref<string[]>([]),
		scope = effectScope(),
		goTo = (index: number) => {
			activeSlideIndex.value = Math.max(0, Math.min(index, slideCount - 1));
		};
	const result = scope.run(() =>
		useSlideChangeSelection({ activeSlideIndex, goTo, selectedElementIds }),
	)!;
	return { ...result, activeSlideIndex, selectedElementIds, goTo, stop: () => scope.stop() };
}

describe('useSlideChangeSelection', () => {
	it('keeps a selection made with its slide change, and clears on every other one', async () => {
		const s = setup();
		s.showSlideSelection(2, ['c']);
		await nextTick();
		expect(s.activeSlideIndex.value).toBe(2);
		expect(s.selectedElementIds.value).toStrictEqual(['c']);
		s.goTo(1);
		await nextTick();
		expect(s.selectedElementIds.value).toStrictEqual([]);
		s.stop();
	});

	it('keeps the selection for one slide change only', async () => {
		const s = setup();
		s.showSlideSelection(1, ['b']);
		await nextTick();
		s.selectedElementIds.value = ['a'];
		s.goTo(0);
		await nextTick();
		expect(s.selectedElementIds.value).toStrictEqual([]);
		// Back to the slide the selection was once made for: that note is gone.
		s.selectedElementIds.value = ['a'];
		s.goTo(1);
		await nextTick();
		expect(s.selectedElementIds.value).toStrictEqual([]);
		// Found on the slide already showing: no slide change, so the next one still clears.
		s.showSlideSelection(1, ['a']);
		await nextTick();
		expect(s.selectedElementIds.value).toStrictEqual(['a']);
		s.goTo(2);
		await nextTick();
		expect(s.selectedElementIds.value).toStrictEqual([]);
		s.stop();
	});

	it('notes the slide it actually landed on', async () => {
		const s = setup(2);
		s.showSlideSelection(5, ['b']);
		await nextTick();
		expect(s.activeSlideIndex.value).toBe(1);
		expect(s.selectedElementIds.value).toStrictEqual(['b']);
		s.stop();
	});
});
