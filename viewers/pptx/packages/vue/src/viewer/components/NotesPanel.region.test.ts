import { flushPromises, mount } from '@vue/test-utils';
import { resolveCustomization } from 'ooxml-ui/pptx';
import type { ViewerCustomization } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';
import { computed, shallowRef } from 'vue';
import type { Component } from 'vue';

import { ViewerCustomizationKey } from '../composables/useViewerCustomization';
import { ViewerRegionBaseKey } from '../composables/useViewerRegions';
import MobileSlidesSheet from './MobileSlidesSheet.vue';
import NotesPanel from './NotesPanel.vue';

function provideAll(customization: ViewerCustomization): Record<symbol, unknown> {
	return {
		[ViewerCustomizationKey as symbol]: shallowRef(resolveCustomization(customization)),
		[ViewerRegionBaseKey as symbol]: computed(() => ({
			viewer: {},
			mode: 'edit',
			canEdit: true,
			activeSlideIndex: 0,
			slideCount: 1,
			selectedElementIds: [],
			locale: 'en',
		})),
	};
}

const slide = { id: 's1', rId: 'r1', slideNumber: 1, elements: [] };

describe('notes region', () => {
	it('replaces the whole panel including the header and stays mounted when collapsed', async () => {
		const destroy = vi.fn();
		const render = vi.fn((container: HTMLElement) => {
			container.textContent = 'host notes';
			return { update: vi.fn(), destroy };
		});
		const wrapper = mount(NotesPanel, {
			props: { slide, expanded: true },
			global: { provide: provideAll({ regions: { notes: render } }) },
		});
		await flushPromises();
		expect(wrapper.get('[data-pptx-region="notes"]').text()).toBe('host notes');
		expect(wrapper.find('[data-pptx-chrome="notes-header"]').exists()).toBeFalsy();
		expect(wrapper.find('.pptx-vue-notes-rich').exists()).toBeFalsy();

		await wrapper.setProps({ expanded: false });
		expect((wrapper.element as HTMLElement).style.display).toBe('none');
		expect(wrapper.find('[data-pptx-region="notes"]').exists()).toBeTruthy();
		expect(render).toHaveBeenCalledOnce();
		expect(destroy).not.toHaveBeenCalled();
		wrapper.unmount();
		expect(destroy).toHaveBeenCalledOnce();
	});

	it('replaces the mobile (embedded) notes sheet content too', async () => {
		const render = vi.fn((container: HTMLElement) => {
			container.textContent = 'host notes';
		});
		const wrapper = mount(NotesPanel, {
			props: { slide, expanded: true, embedded: true },
			global: { provide: provideAll({ regions: { notes: render } }) },
		});
		await flushPromises();
		expect(wrapper.get('[data-pptx-region="notes"]').text()).toBe('host notes');
		expect(wrapper.find('.pptx-vue-notes-rich').exists()).toBeFalsy();
	});

	it('replaces the built-in thumbnails in the mobile slides sheet', async () => {
		const render = vi.fn((container: HTMLElement) => {
			container.textContent = 'host slides';
		});
		const wrapper = mount(MobileSlidesSheet as Component, {
			props: {
				open: true,
				slides: [slide],
				activeIndex: 0,
				canvasSize: { width: 960, height: 540 },
				mediaDataUrls: new Map(),
				canEdit: true,
			},
			global: { provide: provideAll({ regions: { slidesPane: render } }) },
			attachTo: document.body,
		});
		await flushPromises();
		expect(document.body.querySelector('[data-pptx-region="slidesPane"]')?.textContent).toBe(
			'host slides',
		);
		wrapper.unmount();
	});
});
