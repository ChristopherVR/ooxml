import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

/**
 * The find bar through the whole viewer: the viewer drops the selection whenever the active
 * slide changes, so a match on another slide is only selected if the bar selects it after that.
 */
function text(value: string): PptxElement {
	return {
		type: 'text',
		id: value,
		x: 10,
		y: 10,
		width: 200,
		height: 40,
		rotation: 0,
		text: value,
		textSegments: [{ text: value, style: {} }],
	} as PptxElement;
}

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
	document.body.innerHTML = '';
});

async function viewerWithTwoSlides(): Promise<PowerPointViewerExpose> {
	mounted = mount(PowerPointViewer, {
		props: { content: null, canEdit: true },
		attachTo: document.body,
	});
	await flushPromises();
	const viewer = mounted.vm as unknown as PowerPointViewerExpose;
	while (viewer.getSlides().length < 2) {
		viewer.addSlide();
	}
	viewer.setActiveSlideIndex(0);
	await flushPromises();
	viewer.addElement(text('quarter one'));
	viewer.setActiveSlideIndex(1);
	await flushPromises();
	viewer.addElement(text('quarter two'));
	viewer.setActiveSlideIndex(0);
	await flushPromises();
	return viewer;
}

const idWith = (viewer: PowerPointViewerExpose, slide: number, needle: string) =>
	viewer.getSlides()[slide]?.elements.find((e) => (e as { text?: string }).text === needle)?.id;

describe('powerPointViewer find and replace', () => {
	it('selects the matched element, also on another slide, and closing clears the count', async () => {
		const viewer = await viewerWithTwoSlides();
		(mounted!.vm as unknown as { findOpen: boolean }).findOpen = true;
		await flushPromises();
		await mounted!.get('input[aria-label="Search text"]').setValue('quarter');
		await flushPromises();
		expect(viewer.getSelectedElementIds()).toStrictEqual([idWith(viewer, 0, 'quarter one')]);

		await mounted!.get('button[aria-label="Next match"]').trigger('click');
		await flushPromises();
		expect(mounted!.get('.pptx-vue-fr-counter').text()).toBe('2 / 2');
		expect(viewer.getActiveSlideIndex()).toBe(1);
		expect(viewer.getSelectedElementIds()).toStrictEqual([idWith(viewer, 1, 'quarter two')]);

		(mounted!.vm as unknown as { findOpen: boolean }).findOpen = false;
		await flushPromises();
		(mounted!.vm as unknown as { findOpen: boolean }).findOpen = true;
		await flushPromises();
		// Reopened, the kept query is searched again from the first match.
		expect(mounted!.get('.pptx-vue-fr-counter').text()).toBe('1 / 2');
	});
});
