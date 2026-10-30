import { mount } from '@vue/test-utils';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import CustomShowsDialog from './CustomShowsDialog.vue';
import { footerAction } from './dialog-footer.test-support';

const slides = [1, 2, 3].map(
	(n) =>
		({
			id: `s${n}`,
			rId: `rId${n}`,
			slideNumber: n,
			elements:
				n === 1
					? [{ id: 'e1', type: 'shape', actionClick: { action: 'ppaction://customshow?id=7' } }]
					: [],
		}) as unknown as PptxSlide,
);
const shows: PptxCustomShow[] = [
	{ id: '7', name: 'Gone', slideRIds: [] },
	{ id: '8', name: 'Kept', slideRIds: ['rId1', 'rId2'] },
];

const wrappers: ReturnType<typeof mount>[] = [];
afterEach(() => {
	for (const wrapper of wrappers.splice(0)) {
		wrapper.unmount();
	}
});

function open(activeShowId: string | null) {
	const wrapper = mount(CustomShowsDialog, {
		props: { open: true, customShows: shows, slides, activeShowId },
		global: { stubs: { teleport: true } },
	});
	wrappers.push(wrapper);
	return wrapper;
}

describe('customShowsDialog', () => {
	it('copies a show, moves its slides and clears jumps into a removed show', async () => {
		const wrapper = open('8');
		const click = async (label: string, scope = '.pptx-vue-cs-dialog') => {
			await wrapper
				.findAll(`${scope} button`)
				.find((button) => button.text() === label)!
				.trigger('click');
		};
		await click('Copy', '.pptx-vue-cs-shows');
		await wrapper.findAll('[aria-label="Slides in show"] li')[1]!.trigger('click');
		await click('Move up', '.pptx-vue-cs-lists');
		await wrapper.findAll('[aria-label="Saved Shows"] li')[0]!.trigger('click');
		await click('Remove', '.pptx-vue-cs-shows');
		footerAction(wrapper.element, 'save')!.click();
		await wrapper.vm.$nextTick();

		const [savedShows, cleared] = wrapper.emitted('save')![0] as [
			PptxCustomShow[],
			PptxSlide[] | undefined,
		];
		expect(savedShows.map((show) => show.name)).toEqual(['Kept', 'Copy of Kept']);
		expect(savedShows[1]?.slideRIds).toEqual(['rId2', 'rId1']);
		expect(cleared?.[0]?.elements[0]).not.toHaveProperty('actionClick');
		expect(wrapper.emitted('close')).toHaveLength(1);
	});

	it('lets the user restrict playback to a show or lift the restriction', async () => {
		const wrapper = open(null);
		const select = wrapper.find('.pptx-vue-cs-select');
		(select.element as HTMLSelectElement).value = '8';
		await select.trigger('change');
		expect(wrapper.emitted('select')).toEqual([['8']]);
	});
});
