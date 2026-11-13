import { mount } from '@vue/test-utils';
import type { PptxCustomShow, PptxPresentationProperties } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { footerAction } from './dialog-footer.test-support';
import SetUpSlideShowDialog from './SetUpSlideShowDialog.vue';

const SHOWS: PptxCustomShow[] = [{ id: '4', name: 'Short', slideRIds: ['rId1'] }];
const wrappers: ReturnType<typeof mount>[] = [];

afterEach(() => {
	for (const wrapper of wrappers.splice(0)) {
		wrapper.unmount();
	}
});

function saved(act: (radios: HTMLElement[]) => void): PptxPresentationProperties {
	const wrapper = mount(SetUpSlideShowDialog, {
		props: { open: true, properties: {}, customShows: SHOWS, slideCount: 6 },
		global: { stubs: { teleport: true } },
		attachTo: document.body,
	});
	wrappers.push(wrapper);
	// Order: show type x3, all slides, from/to, custom show, advance x2.
	const radios = Array.from(wrapper.element.querySelectorAll<HTMLElement>('pptx-ui-radio'));
	act(radios);
	footerAction(wrapper.element, 'ok')?.click();
	const events = wrapper.emitted('save');
	expect(events).toHaveLength(1);
	return (events?.[0] as [PptxPresentationProperties])[0];
}

const change = (radio: HTMLElement | undefined) => {
	radio?.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('vue set up slide show dialog follows the shared rules', () => {
	it('a kiosk show loops continuously', () => {
		expect(saved((radios) => change(radios[2]))).toMatchObject({
			showType: 'kiosk',
			loopContinuously: true,
		});
	});

	it('a range starts as the whole deck', () => {
		expect(saved((radios) => change(radios[4]))).toMatchObject({
			showSlidesMode: 'range',
			showSlidesFrom: 1,
			showSlidesTo: 6,
		});
	});

	it('a custom show defaults to the first show', () => {
		expect(saved((radios) => change(radios[5]))).toMatchObject({
			showSlidesMode: 'customShow',
			showSlidesCustomShowId: '4',
		});
	});
});
