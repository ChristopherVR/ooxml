// oxlint-disable react-hooks/rules-of-hooks -- Vue composables run in `setup`, not React hooks.
/**
 * `SlidesPane`, the slide thumbnail rail as a standalone building block: a host
 * places it beside its own `SlideCanvas` and drives it from `useLoadContent`
 * and the slide and section operations, without `PowerPointViewer`.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { flushPromises, mount } from '@vue/test-utils';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { defineComponent, h, ref } from 'vue';

import { useEditorHistory } from '../composables/useEditorHistory';
import { useLoadContent } from '../composables/useLoadContent';
import type { SectionGroup } from '../composables/useSectionOperations';
import { useSlideOperations } from '../composables/useSlideOperations';
import ContextMenu from './ContextMenu.vue';
import SlidesPane from './SlidesPane.vue';

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx');
const canvasSize = { width: 960, height: 540 };

function slide(id: string): PptxSlide {
	return { id, rId: '', slideNumber: 1, elements: [] } as unknown as PptxSlide;
}

/** A host shell: the rail driven by the same slide operations a custom editor holds. */
function mountHost(options: { canEdit?: boolean; sectionGroups?: SectionGroup[] } = {}) {
	let api!: {
		slides: ReturnType<typeof ref<PptxSlide[]>>;
		active: ReturnType<typeof ref<number>>;
		history: ReturnType<typeof useEditorHistory>;
	};
	const Host = defineComponent({
		setup() {
			const slides = ref<PptxSlide[]>([slide('a'), slide('b'), slide('c')]);
			const active = ref(0);
			const history = useEditorHistory(slides);
			const ops = useSlideOperations({
				slides,
				activeSlideIndex: active,
				pushHistory: history.pushHistory,
			});
			api = { slides, active, history } as typeof api;
			return () =>
				h(SlidesPane, {
					slides: slides.value,
					activeSlideIndex: active.value,
					canvasSize,
					canEdit: options.canEdit ?? true,
					sectionGroups: options.sectionGroups ?? [],
					onSelect: (index: number) => {
						active.value = index;
					},
					onAddSlide: () => ops.addSlide(),
					onDuplicateSlide: ops.duplicateSlide,
					onDeleteSlide: ops.deleteSlide,
				});
		},
	});
	const wrapper = mount(Host);
	return { wrapper, api: () => api };
}

const rows = (wrapper: ReturnType<typeof mount>) =>
	wrapper.findAll('[data-pptx-chrome="slide-row"]');

async function chooseFromMenu(
	wrapper: ReturnType<typeof mount>,
	rowIndex: number,
	command: string,
): Promise<void> {
	const thumbs = wrapper.findAll('[aria-label^="Go to slide "]');
	await thumbs[rowIndex].trigger('contextmenu', { clientX: 40, clientY: 60 });
	const menu = wrapper.findAllComponents(ContextMenu).find((m) => m.props('open'));
	expect(menu).toBeDefined();
	menu?.vm.$emit('select', command);
	await flushPromises();
}

describe('slidesPane building block', () => {
	it('is exported from the package root and the /viewer subpath', async () => {
		const root = await import('../../index');
		const viewer = await import('../index');
		expect(root.SlidesPane).toBe(SlidesPane);
		expect(viewer.SlidesPane).toBe(SlidesPane);
		expect(viewer.useSlideOperations).toBe(useSlideOperations);
		expect(viewer.useSectionOperations).toBeTypeOf('function');
	});

	it('numbers thumbnails from the deck start number (firstSlideNum="45")', async () => {
		const Host = defineComponent({
			setup() {
				const content = useLoadContent(new Uint8Array(readFileSync(FIXTURE)));
				return () =>
					h(SlidesPane, {
						slides: content.slides.value,
						templateElementsBySlideId: content.templateElementsBySlideId.value,
						activeSlideIndex: 0,
						canvasSize: content.canvasSize.value,
						mediaDataUrls: content.mediaDataUrls.value,
						headerFooter: content.headerFooter.value,
						customProperties: content.customProperties.value,
						firstSlideNumber: content.firstSlideNumber.value,
						theme: content.theme.value,
						tableStyleMap: content.tableStyleMap.value,
					});
			},
		});
		const wrapper = mount(Host);
		try {
			await vi.waitFor(
				async () => {
					await flushPromises();
					expect(wrapper.findAll('[data-pptx-chrome="slide-frame"]')).toHaveLength(3);
				},
				{ timeout: 15000 },
			);
			const frames = wrapper.findAll('[data-pptx-chrome="slide-frame"]').map((f) => f.text());
			expect(frames[0]).toContain('Slide 45');
			expect(frames[1]).toContain('Slide 46');
			expect(frames[2]).toContain('Slide 47');
			expect(frames.join(' ')).not.toMatch(/Slide [123]\b/u);
		} finally {
			wrapper.unmount();
		}
	});

	it('emits the shared DOM markers: rows, aria-current, frames and the footer', () => {
		const { wrapper } = mountHost();
		expect(rows(wrapper)).toHaveLength(3);
		expect(rows(wrapper)[0].attributes('aria-current')).toBe('true');
		expect(wrapper.findAll('[data-pptx-chrome="slide-frame"]')).toHaveLength(3);
		expect(wrapper.findAll('[data-pptx-chrome="slide-footer"] button')).toHaveLength(1);
	});

	it('selecting a thumbnail moves the active slide without an undo step or a deck change', async () => {
		const { wrapper, api } = mountHost();
		const before = api().slides.value;
		await rows(wrapper)[2].trigger('click');
		expect(api().active.value).toBe(2);
		expect(rows(wrapper)[2].attributes('aria-current')).toBe('true');
		expect(api().slides.value).toBe(before);
		expect(api().history.canUndo.value).toBeFalsy();
	});

	it('the Layout entry makes the slide active, then asks the host for the gallery', async () => {
		const wrapper = mount(SlidesPane, {
			props: { slides: [slide('a'), slide('b')], activeSlideIndex: 0, canvasSize, canEdit: true },
		});
		await chooseFromMenu(wrapper, 1, 'layout');
		expect(wrapper.emitted('select')).toStrictEqual([[1]]);
		expect(wrapper.emitted('open-layout-gallery')).toStrictEqual([
			[{ slideIndex: 1, x: 40, y: 60 }],
		]);
	});

	it('duplicates every selected slide, each copy right after its source', async () => {
		const { wrapper, api } = mountHost();
		await rows(wrapper)[0].trigger('click');
		await rows(wrapper)[2].trigger('click', { ctrlKey: true });
		await chooseFromMenu(wrapper, 2, 'duplicate');
		const ids = api().slides.value.map((s) => s.id);
		expect(ids).toHaveLength(5);
		expect([ids[0], ids[2], ids[3]]).toStrictEqual(['a', 'b', 'c']);
		expect(ids[1]).not.toBe('a');
		expect(ids[4]).not.toBe('c');
	});

	it('deletes every selected slide', async () => {
		const { wrapper, api } = mountHost();
		await rows(wrapper)[0].trigger('click');
		await rows(wrapper)[2].trigger('click', { ctrlKey: true });
		await chooseFromMenu(wrapper, 2, 'delete');
		expect(api().slides.value.map((s) => s.id)).toStrictEqual(['b']);
	});

	it('adds a slide from the footer through the host operation', async () => {
		const { wrapper, api } = mountHost();
		await wrapper.get('[data-pptx-chrome="slide-footer"] button').trigger('click');
		expect(api().slides.value).toHaveLength(4);
		expect(rows(wrapper)).toHaveLength(4);
		expect(rows(wrapper)[1].attributes('aria-current')).toBe('true');
	});

	it('shows section headers and maps section events', async () => {
		const slides = [slide('a'), slide('b')];
		const wrapper = mount(SlidesPane, {
			props: {
				slides,
				activeSlideIndex: 0,
				canvasSize,
				canEdit: true,
				sectionGroups: [
					{ section: { id: 'sec1', name: 'Intro', slideIds: [] }, slides, slideIndexes: [0, 1] },
				] as unknown as SectionGroup[],
			},
		});
		expect(wrapper.get('[data-pptx-chrome="section-header"]').text()).toContain('Intro');
		expect(wrapper.findAll('[aria-label^="Go to slide "]')).toHaveLength(2);
		await chooseFromMenu(wrapper, 1, 'layout');
		expect(wrapper.emitted('select')).toStrictEqual([[1]]);
		expect(wrapper.emitted('open-layout-gallery')?.[0]?.[0]).toMatchObject({ slideIndex: 1 });
		await chooseFromMenu(wrapper, 0, 'add-section');
		expect(wrapper.emitted('add-section')?.[0]?.[1]).toBe(0);
		expect(wrapper.emitted('add-section')?.[0]?.[0]).toBeTypeOf('string');
	});

	for (const sectioned of [false, true]) {
		it(`shows no footer when read-only (${sectioned ? 'sectioned' : 'flat'})`, () => {
			const slides = [slide('a')];
			const wrapper = mount(SlidesPane, {
				props: {
					slides,
					activeSlideIndex: 0,
					canvasSize,
					canEdit: false,
					sectionGroups: sectioned
						? ([
								{ section: { id: 's', name: 'S', slideIds: [] }, slides, slideIndexes: [0] },
							] as unknown as SectionGroup[])
						: [],
				},
			});
			expect(wrapper.find('[data-pptx-chrome="slide-footer"]').exists()).toBeFalsy();
		});
	}
});
