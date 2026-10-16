import { mount } from '@vue/test-utils';
import type { RibbonGalleryApplyResult } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { computed, ref, shallowRef } from 'vue';

import { useDeckStyleHost } from '../../composables/useDeckStyleHost';
import { RibbonGalleryHostKey } from '../../composables/useRibbonGalleryHost';
import type { RibbonGalleryHost } from '../../composables/useRibbonGalleryHost';
import RibbonGallery from './RibbonGallery.vue';

/**
 * The Design galleries: hover preview, the tile context menu and the footer command reach the
 * gallery host, and the deck-style host swaps the slides in and out without recording history.
 */
beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

const slideOf = (color: string): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements: [], backgroundColor: color }) as PptxSlide;

function deckRefs() {
	return {
		slides: ref<PptxSlide[]>([slideOf('#FFFFFF')]),
		templateElementsBySlideId: shallowRef<Record<string, PptxElement[]>>({}),
		pptxTheme: ref<PptxTheme | undefined>({ name: 'Office' }),
		themeColorMap: ref<Record<string, string> | undefined>(undefined),
		held: ref(false),
	};
}

describe('useDeckStyleHost', () => {
	it('swaps the preview in and the original back, holding autosave and collaboration', () => {
		const refs = deckRefs();
		const pushHistory = vi.fn();
		const heldDuring: boolean[] = [];
		const host = useDeckStyleHost({
			...refs,
			handler: ref({
				applyBackgroundStyleToSlides: (slides: PptxSlide[]) => ({
					slides: slides.map((s) => ({ ...s, backgroundColor: '#000000' })),
					templateElementsBySlideId: {},
				}),
			} as never),
			selectedSlideIds: () => ['s1'],
			pushHistory,
		});
		const original = refs.slides.value;
		host.preview({ kind: 'backgroundStyle', styleIndex: 4, scope: 'all' });
		heldDuring.push(refs.held.value);
		expect(refs.slides.value[0]!.backgroundColor).toBe('#000000');
		expect(host.held()?.slides).toBe(original);
		host.endPreview();
		heldDuring.push(refs.held.value);
		expect(refs.slides.value).toBe(original);
		expect(heldDuring).toEqual([true, false]);
		expect(pushHistory).not.toHaveBeenCalled();
	});

	it('records one undo step when a pick commits, from the real slides', async () => {
		const refs = deckRefs();
		const order: string[] = [];
		const commitMaster = vi.fn(async () => true);
		const host = useDeckStyleHost({
			...refs,
			handler: ref({
				applyBackgroundStyleToSlides: (slides: PptxSlide[]) => ({
					slides: slides.map((s) => ({ ...s, backgroundColor: '#000000' })),
					templateElementsBySlideId: {},
				}),
				commitMasterBackgroundStyle: commitMaster,
			} as never),
			selectedSlideIds: () => ['s1'],
			pushHistory: () => order.push(`push ${refs.slides.value[0]!.backgroundColor}`),
		});
		host.preview({ kind: 'backgroundStyle', styleIndex: 4, scope: 'all' });
		await host.commit({ kind: 'backgroundStyle', styleIndex: 4, scope: 'all' });
		expect(order).toEqual(['push #FFFFFF']);
		expect(refs.slides.value[0]!.backgroundColor).toBe('#000000');
		expect(commitMaster).toHaveBeenCalledWith(4);
		expect(refs.held.value).toBe(false);
	});
});

describe('RibbonGallery (Background Styles)', () => {
	function mountGallery() {
		const dispatched: RibbonGalleryApplyResult[] = [];
		const previews: Array<RibbonGalleryApplyResult | null> = [];
		const ended = vi.fn();
		const host: RibbonGalleryHost = {
			context: computed(() => ({ element: null, theme: { name: 'Office' } })),
			dispatch: (result) => dispatched.push(result),
			preview: (result) => previews.push(result),
			endPreview: ended,
		};
		const wrapper = mount(RibbonGallery, {
			props: {
				gallery: 'themeBackgrounds',
				control: 'design.variants.backgroundStyles',
				mode: 'dropdown',
			},
			global: { provide: { [RibbonGalleryHostKey as symbol]: host } },
			attachTo: document.body,
		});
		wrapper.element.querySelector<HTMLButtonElement>('[data-ribbon-gallery]')!.click();
		return { wrapper, dispatched, previews, ended };
	}

	it('previews a hovered tile and ends the preview', () => {
		const { wrapper, previews, ended } = mountGallery();
		const tile = wrapper.element.querySelector<HTMLElement>('[data-gallery-item="5"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		expect(previews).toEqual([{ kind: 'backgroundStyle', styleIndex: 5, scope: 'all' }]);
		tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(100);
		expect(ended).toHaveBeenCalled();
		wrapper.unmount();
	});

	it('applies to the selected slides from the tile menu', () => {
		const { wrapper, dispatched } = mountGallery();
		const tile = wrapper.element.querySelector<HTMLElement>('[data-gallery-item="2"]')!;
		tile.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
		const entries = [...wrapper.element.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')];
		entries[1]!.click();
		expect(dispatched).toEqual([{ kind: 'backgroundStyle', styleIndex: 2, scope: 'selected' }]);
		wrapper.unmount();
	});

	it('emits the Format Background command from the footer', () => {
		const { wrapper, dispatched } = mountGallery();
		wrapper.element
			.querySelector<HTMLButtonElement>('[data-gallery-footer="formatBackground"]')!
			.click();
		expect(wrapper.emitted('command')).toEqual([['formatBackground']]);
		expect(dispatched).toEqual([]);
		wrapper.unmount();
	});
});
