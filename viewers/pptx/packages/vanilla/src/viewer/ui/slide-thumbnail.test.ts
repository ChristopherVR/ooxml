// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

// Import through the package entry so the test also covers the public export.
import { renderSlideThumbnail } from '../../index';

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
const render = (elements: PptxElement[], templateElements: PptxElement[] = []) =>
	renderSlideThumbnail({
		document,
		slide: slide(elements),
		templateElements,
		canvasSize,
		previewWidth: 120,
	});
const stage = (box: HTMLElement) => box.firstElementChild as HTMLElement | null;
const drawn = (box: HTMLElement) => [...(stage(box)?.children ?? [])] as HTMLElement[];

describe('renderSlideThumbnail', () => {
	it('draws the slide stage scaled to the width the host picks, at the deck aspect ratio', () => {
		const box = render([shape('a')]);
		expect(box.style.height).toBe('67.5px');
		expect(box.style.width).toBe('100%');
		expect(box.style.overflow).toBe('hidden');
		expect(stage(box)?.style.width).toBe('960px');
		expect(stage(box)?.style.transform).toBe('scale(0.125)');
	});

	it('draws inherited template elements beneath the slide elements', () => {
		const box = render([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn(box).map((node) => node.style.left)).toStrictEqual(['10px', '200px']);
	});

	it('stays static: inert, no element ids and no editing contract', () => {
		const box = render([shape('a')], [shape('layout', 10)]);
		expect(stage(box)?.inert).toBeTruthy();
		expect(box.querySelector('[data-element-id]')).toBeNull();
		expect(box.querySelector('[data-pptx-element]')).toBeNull();
		expect(box.querySelector('[aria-roledescription="slide"]')).toBeNull();
		expect(box.querySelector('[data-pptx-chrome]')).toBeNull();
	});

	it('hides template elements when the slide hides background graphics', () => {
		const box = renderSlideThumbnail({
			document,
			slide: { ...slide([shape('own', 200)]), showMasterShapes: false },
			templateElements: [shape('layout', 10)],
			canvasSize,
			previewWidth: 120,
		});
		expect(drawn(box).map((node) => node.style.left)).toStrictEqual(['200px']);
	});

	it('follows edits to the slide model without a save', () => {
		expect(drawn(render([shape('a')]))).toHaveLength(1);
		expect(drawn(render([shape('a'), shape('b', 300)]))).toHaveLength(2);
	});
});
