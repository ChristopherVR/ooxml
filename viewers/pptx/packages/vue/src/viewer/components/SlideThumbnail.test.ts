import { mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

// Import through the package entry so the test also covers the public export.
import { SlideThumbnail } from '../../index';

function shape(id: string, x = 0): PptxElement {
	return {
		id,
		type: 'shape',
		shapeType: 'rect',
		x,
		y: 0,
		width: 100,
		height: 50,
		shapeStyle: { fillColor: '#ff0000' },
	} as unknown as PptxElement;
}

function slide(elements: PptxElement[]): PptxSlide {
	return {
		id: 'ppt/slides/slide1.xml',
		rId: 'rId2',
		slideNumber: 1,
		elements,
		showMasterShapes: true,
	} as unknown as PptxSlide;
}

const canvasSize = { width: 960, height: 540 };

function mountThumbnail(elements: PptxElement[], templateElements: PptxElement[] = []) {
	return mount(SlideThumbnail, {
		props: { slide: slide(elements), templateElements, canvasSize, previewWidth: 120 },
	});
}

const stage = (wrapper: ReturnType<typeof mountThumbnail>) =>
	wrapper.element.firstElementChild as HTMLElement;
// Thumbnails draw elements without their ids, so ids stay unique to the editing canvas.
const drawn = (wrapper: ReturnType<typeof mountThumbnail>) => [
	...stage(wrapper).querySelectorAll<HTMLElement>(':scope > *'),
];

describe('slideThumbnail', () => {
	it('draws the slide stage scaled to the width the host picks, at the deck aspect ratio', () => {
		const wrapper = mountThumbnail([shape('a')]);
		const box = wrapper.element as HTMLElement;
		expect(box.style.height).toBe('67.5px');
		expect(box.style.overflow).toBe('hidden');
		expect(stage(wrapper).style.width).toBe('960px');
		expect(stage(wrapper).style.transform).toBe('scale(0.125)');
		wrapper.unmount();
	});

	it('draws inherited template elements beneath the slide elements', () => {
		const wrapper = mountThumbnail([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn(wrapper).map((node) => node.style.left)).toStrictEqual(['10px', '200px']);
		wrapper.unmount();
	});

	it('stays static: no element ids and no interaction hooks', () => {
		const wrapper = mountThumbnail([shape('a')], [shape('layout')]);
		expect(wrapper.element.querySelector('[data-element-id]')).toBeNull();
		expect(wrapper.element.querySelector('[data-pptx-element]')).toBeNull();
		expect(stage(wrapper).getAttribute('aria-hidden')).toBe('true');
		wrapper.unmount();
	});

	it('follows edits to the slide model without a save', async () => {
		const wrapper = mountThumbnail([shape('a')]);
		expect(drawn(wrapper)).toHaveLength(1);
		await wrapper.setProps({ slide: slide([shape('a'), shape('b')]) });
		expect(drawn(wrapper)).toHaveLength(2);
		wrapper.unmount();
	});

	it('follows slide metadata edits that keep the same elements', async () => {
		const number = {
			...shape('number'),
			type: 'text',
			textSegments: [{ text: '#', style: {}, fieldType: 'slidenum' }],
		} as unknown as PptxElement;
		const base = slide([number]);
		const wrapper = mount(SlideThumbnail, {
			props: { slide: base, canvasSize, previewWidth: 120 },
		});
		expect(wrapper.text()).toBe('1');
		await wrapper.setProps({ slide: { ...base, slideNumber: 7 } });
		expect(wrapper.text()).toBe('7');
		wrapper.unmount();
	});

	// The same badge, from the same shared descriptor, in every binding.
	it('draws the transition badge in the corner, at the slide scale', async () => {
		const wrapper = mount(SlideThumbnail, {
			props: {
				slide: { ...slide([shape('a')]), transition: { type: 'fade' } } as PptxSlide,
				canvasSize,
				previewWidth: 120,
			},
		});
		const badge = () =>
			wrapper.element.querySelector<HTMLElement>(':scope > [data-pptx-thumbnail-transition]');
		expect(badge()?.textContent?.trim()).toBe('Fade');
		expect(badge()?.title).toBe('Transition: Fade');
		expect(badge()?.style.transform).toBe('scale(0.125)');
		// The drawn slide's corner, 120px from the box's left edge.
		expect(badge()?.style.right).toBe('calc(100% - 119.75px)');
		await wrapper.setProps({ slide: { ...slide([shape('a')]), transition: { type: 'cut' } } });
		expect(badge()).toBeNull();
		await wrapper.setProps({ slide: slide([shape('a')]) });
		expect(badge()).toBeNull();
		wrapper.unmount();
	});

	it('drops the template layer when the slide hides background graphics', async () => {
		const wrapper = mountThumbnail([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn(wrapper)).toHaveLength(2);
		await wrapper.setProps({ slide: { ...slide([shape('own', 200)]), showMasterShapes: false } });
		expect(drawn(wrapper).map((node) => node.style.left)).toStrictEqual(['200px']);
		wrapper.unmount();
	});
});
