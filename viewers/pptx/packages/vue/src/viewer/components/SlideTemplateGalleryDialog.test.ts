import { mount } from '@vue/test-utils';
import { SLIDE_TEMPLATES } from 'ooxml-ui/pptx';
import type { SlideTemplateId } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { allButtons } from './dialog-footer.test-support';
import SlidesGroup from './ribbon/SlidesGroup.vue';
import SlideTemplateGalleryDialog from './SlideTemplateGalleryDialog.vue';

afterEach(() => {
	document.body.innerHTML = '';
});

function dialogEl(): HTMLElement {
	const el = document.body.querySelector<HTMLElement>('[role="dialog"]:not([hidden])');
	if (!el) {
		throw new Error('dialog not found');
	}
	return el;
}

function optionTiles(): HTMLButtonElement[] {
	return Array.from(document.body.querySelectorAll<HTMLButtonElement>('[role="option"]'));
}

function tileByLabel(label: string): HTMLButtonElement {
	const tile = optionTiles().find((b) => b.getAttribute('aria-label') === label);
	if (!tile) {
		throw new Error(`option "${label}" not found`);
	}
	return tile;
}

function footerButton(label: string): HTMLButtonElement {
	const btn = allButtons(document.body).find((b) => b.textContent?.trim() === label);
	if (!btn) {
		throw new Error(`button "${label}" not found`);
	}
	return btn;
}

describe('slideTemplateGalleryDialog', () => {
	it('renders an accessible dialog with one option per catalog template', () => {
		mount(SlideTemplateGalleryDialog, {
			props: { open: true },
			attachTo: document.body,
		});

		expect(dialogEl().getAttribute('aria-label')).toBe('Slide Templates');
		const listbox = document.body.querySelector('[role="listbox"]');
		expect(listbox?.getAttribute('aria-label')).toBe('Slide template gallery');

		const tiles = optionTiles();
		expect(tiles).toHaveLength(SLIDE_TEMPLATES.length);
		expect(tiles.map((tile) => tile.getAttribute('aria-label'))).toContain('Title Slide');
		expect(tiles.map((tile) => tile.getAttribute('aria-label'))).toContain('Agenda');
		expect(tileByLabel('Title Slide').getAttribute('title')).toBe(
			'Large title with subtitle and accent bar',
		);
		expect(tiles.every((tile) => tile.getAttribute('aria-selected') === 'false')).toBeTruthy();
	});

	it('keeps Insert disabled until a template is selected, then inserts it', async () => {
		const wrapper = mount(SlideTemplateGalleryDialog, {
			props: { open: true },
			attachTo: document.body,
		});

		expect(footerButton('Insert').disabled).toBeTruthy();

		tileByLabel('Agenda').click();
		await wrapper.vm.$nextTick();
		expect(tileByLabel('Agenda').getAttribute('aria-selected')).toBe('true');
		expect(footerButton('Insert').disabled).toBeFalsy();

		footerButton('Insert').click();
		await wrapper.vm.$nextTick();

		expect(wrapper.emitted('insert')).toStrictEqual([['agenda']]);
		expect(wrapper.emitted('close')).toHaveLength(1);
	});

	it('inserts immediately on double click', async () => {
		const wrapper = mount(SlideTemplateGalleryDialog, {
			props: { open: true },
			attachTo: document.body,
		});

		tileByLabel('Title Slide').dispatchEvent(new Event('dblclick'));
		await wrapper.vm.$nextTick();

		expect(wrapper.emitted('insert')).toStrictEqual([['title']]);
		expect(wrapper.emitted('close')).toHaveLength(1);
	});

	it('emits close on Cancel without inserting', async () => {
		const wrapper = mount(SlideTemplateGalleryDialog, {
			props: { open: true },
			attachTo: document.body,
		});

		footerButton('Cancel').click();
		await wrapper.vm.$nextTick();

		expect(wrapper.emitted('insert')).toBeUndefined();
		expect(wrapper.emitted('close')).toHaveLength(1);
	});

	it('renders live previews with the provided theme scheme colour', () => {
		mount(SlideTemplateGalleryDialog, {
			props: { open: true, scheme: { accent1: '#BA0021' } },
			attachTo: document.body,
		});

		const gallery = document.body.querySelector('.pptx-vue-template-gallery');
		expect(gallery?.innerHTML.toLowerCase()).toContain('#ba0021');
	});
});

describe('slidesGroup template affordance', () => {
	const base = {
		canEdit: true,
		layoutOptions: [{ path: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide' }],
		onInsertSlideFromLayout: (): void => {},
	};

	/** PowerPoint has no Slide Templates button on Home: it ends the New Slide menu. */
	function templatesEntry(): HTMLButtonElement | undefined {
		document.body
			.querySelector<HTMLElement>(
				'[data-ribbon-control="home.slides.newSlide"] [data-pptx-chrome="split-caret"]',
			)
			?.click();
		return (
			document.body.querySelector<HTMLButtonElement>(
				'.popup [data-ribbon-control="home.slides.slideTemplates"]',
			) ?? undefined
		);
	}

	it('lists Slide Templates in the New Slide menu and routes inserts to the handler', async () => {
		const onInsertSlideFromTemplate = vi.fn<(templateId: SlideTemplateId) => void>();
		const wrapper = mount(SlidesGroup, {
			props: { ...base, onInsertSlideFromTemplate },
			attachTo: document.body,
		});

		const entry = templatesEntry();
		expect(entry).toBeTruthy();
		expect(entry?.textContent).toContain('Slide Templates');

		entry?.click();
		await wrapper.vm.$nextTick();
		expect(dialogEl().getAttribute('aria-label')).toBe('Slide Templates');

		tileByLabel('Key Metrics').click();
		await wrapper.vm.$nextTick();
		footerButton('Insert').click();
		await wrapper.vm.$nextTick();

		expect(onInsertSlideFromTemplate).toHaveBeenCalledExactlyOnceWith('keyMetrics');
		expect(document.body.querySelector('[role="dialog"]:not([hidden])')).toBeNull();
	});

	it('leaves the entry out when no template handler is provided', () => {
		mount(SlidesGroup, { props: { ...base }, attachTo: document.body });
		expect(templatesEntry()).toBeUndefined();
	});
});
