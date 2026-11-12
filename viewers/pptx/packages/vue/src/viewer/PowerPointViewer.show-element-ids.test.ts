// @vitest-environment happy-dom
/**
 * One marked node per element id while a show runs.
 *
 * Vue presents on a separate full-screen stage and keeps the editor mounted
 * (hidden) underneath, so its refs survive the show. The editing canvas kept
 * drawing the active slide there, with its `data-element-id` markers, so every
 * element on that slide answered from two nodes during the show and in the
 * presenter view, however it was opened (it keeps the show stage under its
 * console). The show stage is the surface the user acts on; the hidden canvas
 * now draws no slide until the show ends.
 */
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PresentationToolbar from './components/PresentationToolbar.vue';
import ModeSwitcher from './components/ribbon/ModeSwitcher.vue';
import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

// There is no IndexedDB here, and the recovery write is fire-and-forget.
// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: () => Promise.resolve(true),
}));

/** Two slides; the first holds one rectangle (the loader assigns its id). */
async function deckWithShape(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		const first = data.slides[0]!;
		first.elements = [
			...first.elements,
			{
				id: 'shape',
				type: 'shape',
				shapeType: 'rect',
				x: 100,
				y: 100,
				width: 200,
				height: 100,
				shapeStyle: { fillColor: '#ff0000' },
			} as unknown as PptxElement,
		];
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
});

async function settle(done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !done(); attempt += 1) {
		await flushPromises();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
}

function markedIds(root: Element): string[] {
	return [...root.querySelectorAll('[data-element-id]')].map(
		(node) => node.getAttribute('data-element-id') ?? '',
	);
}

const duplicates = (ids: readonly string[]) => ids.filter((id, at) => ids.indexOf(id) !== at);

/** Mount the viewer and wait for the editing canvas's markers. */
async function openViewer(): Promise<{
	wrapper: VueWrapper;
	viewer: PowerPointViewerExpose;
	editing: string[];
}> {
	const wrapper = mount(PowerPointViewer, {
		props: { content: await deckWithShape(), canEdit: true, filePath: 'show.pptx' },
		attachTo: document.body,
	});
	mounted = wrapper;
	await settle(() => (wrapper.emitted('load')?.length ?? 0) > 0);
	await settle(() => markedIds(document.body).length > 0);
	// Slide 1's elements, drawn by the editing canvas (the rail draws none).
	const editing = markedIds(document.body);
	expect(editing.length).toBeGreaterThan(0);
	expect(duplicates(editing)).toStrictEqual([]);
	return { wrapper, viewer: wrapper.vm as unknown as PowerPointViewerExpose, editing };
}

/** The editor body: rail, editing canvas and panels, hidden while a show runs. */
const editorBody = () => document.querySelector('[data-pptx-chrome="body"]')!;
const presenterConsole = () => document.querySelector('[data-pptx-presenter-control]');

/** The presenter view's entry points: the show toolbar's toggle and the ribbon's Presenter View. */
const presenterEntries: [
	string,
	(wrapper: VueWrapper, viewer: PowerPointViewerExpose) => Promise<void>,
][] = [
	[
		"the show toolbar's presenter toggle",
		async (wrapper, viewer) => {
			viewer.setMode('present');
			await settle(() => document.querySelector('[data-pptx-presenting]') !== null);
			// The toggle's button is in a shadow root.
			wrapper.findComponent(PresentationToolbar).vm.$emit('toggle-presenter-view');
		},
	],
	[
		"the ribbon's Presenter View",
		async (wrapper) => {
			const enter = wrapper.findComponent(ModeSwitcher).props('onEnterPresenterView');
			expect(enter).toBeTypeOf('function');
			enter!();
		},
	],
];

describe('element ids while a show runs', () => {
	it('come from the show stage only, and from the editing canvas again after it', async () => {
		const { viewer, editing } = await openViewer();

		viewer.setMode('present');
		await settle(() => document.querySelector('[data-pptx-presenting]') !== null);
		const show = document.querySelector('[data-pptx-presenting]')!;
		expect(markedIds(show)).toStrictEqual(editing);
		expect(markedIds(document.body)).toStrictEqual(editing);

		// Back in the editor, the canvas draws the slide with its markers again.
		viewer.setMode('edit');
		await settle(() => markedIds(editorBody()).length > 0);
		expect(markedIds(document.body)).toStrictEqual(editing);
	});

	it.each(presenterEntries)(
		'leave the editing canvas in the presenter view opened from %s',
		async (_entry, openPresenter) => {
			const { wrapper, viewer, editing } = await openViewer();

			await openPresenter(wrapper, viewer);
			await settle(() => presenterConsole() !== null);
			expect(presenterConsole()).not.toBeNull();
			expect(markedIds(editorBody())).toStrictEqual([]);
			expect(duplicates(markedIds(document.body))).toStrictEqual([]);

			viewer.setMode('edit');
			await settle(() => markedIds(editorBody()).length > 0);
			expect(presenterConsole()).toBeNull();
			expect(markedIds(document.body)).toStrictEqual(editing);
		},
	);
});
