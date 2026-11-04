import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

// Import through the package entry so the test also covers the public export.
import { SlideThumbnail } from '../../index';
import type { SlideThumbnailProps } from '../../index';

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
const cleanups: (() => void)[] = [];

afterEach(() => {
	for (const cleanup of cleanups.splice(0)) {
		cleanup();
	}
	document.body.replaceChildren();
});

function mountThumbnail(elements: PptxElement[], templateElements: PptxElement[] = []) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const props = $state<SlideThumbnailProps>({
		slide: slide(elements),
		templateElements,
		canvasSize,
		previewWidth: 120,
	});
	const component = mount(SlideThumbnail, { target, props });
	flushSync();
	cleanups.push(() => unmount(component));
	const box = target.firstElementChild as HTMLElement;
	const stage = () => box.firstElementChild as HTMLElement;
	return {
		props,
		box,
		stage,
		// Thumbnails draw elements without their ids, so ids stay unique to the editing canvas.
		drawn: () => [...stage().querySelectorAll<HTMLElement>(':scope > .pptx-svelte-element')],
	};
}

describe('slideThumbnail', () => {
	it('draws the slide stage scaled to the width the host picks, at the deck aspect ratio', () => {
		const { box, stage } = mountThumbnail([shape('a')]);
		expect(box.style.height).toBe('67.5px');
		expect(box.style.overflow).toBe('hidden');
		expect(stage().style.width).toBe('960px');
		expect(stage().style.transform).toBe('scale(0.125)');
	});

	it('draws inherited template elements beneath the slide elements', () => {
		const { drawn } = mountThumbnail([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn().map((node) => node.style.left)).toStrictEqual(['10px', '200px']);
	});

	it('stays static: no element ids and no interaction hooks', () => {
		const { box, stage } = mountThumbnail([shape('a')], [shape('layout')]);
		expect(box.querySelector('[data-element-id]')).toBeNull();
		expect(box.querySelector('[data-pptx-element]')).toBeNull();
		expect(stage().getAttribute('role')).toBeNull();
	});

	it('follows edits to the slide model without a save', () => {
		const { props, drawn } = mountThumbnail([shape('a')]);
		expect(drawn()).toHaveLength(1);
		props.slide = slide([shape('a'), shape('b')]);
		flushSync();
		expect(drawn()).toHaveLength(2);
	});

	it('drops the template layer when the slide hides background graphics', () => {
		const { props, drawn } = mountThumbnail([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn()).toHaveLength(2);
		props.slide = { ...props.slide, showMasterShapes: false };
		flushSync();
		expect(drawn().map((node) => node.style.left)).toStrictEqual(['200px']);
	});
});
