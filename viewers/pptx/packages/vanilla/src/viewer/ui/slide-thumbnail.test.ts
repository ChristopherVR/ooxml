// @vitest-environment happy-dom
import { ensureBarFacePicturePixelSampled, releasePptxRenderCaches } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

// Import through the package entry so the test also covers the public export.
import { disposePptxRenderStage, renderSlideThumbnail } from '../../index';

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

	// The same badge, from the same shared descriptor, in every binding.
	it('draws the transition badge in the corner, at the slide scale', () => {
		const draw = (transition: PptxSlide['transition']) =>
			renderSlideThumbnail({
				document,
				slide: { ...slide([shape('a')]), transition },
				canvasSize,
				previewWidth: 120,
			}).querySelector<HTMLElement>(':scope > [data-pptx-thumbnail-transition]');
		const badge = draw({ type: 'fade' });
		expect(badge?.textContent).toBe('Fade');
		expect(badge?.title).toBe('Transition: Fade');
		expect(badge?.style.transform).toBe('scale(0.125)');
		// The drawn slide's corner, 120px from the box's left edge.
		expect(badge?.style.right).toBe('calc(100% - 119.75px)');
		expect(draw({ type: 'cut' })).toBeNull();
		expect(draw(undefined)).toBeNull();
	});
});

describe('renderSlideThumbnail cleanup contract', () => {
	const chart = {
		type: 'chart',
		id: 'chart',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
	} as unknown as PptxElement;
	const rows: HTMLElement[] = [];
	let sampleIndex = 0;
	/** Resolve one picture colour sample: every subscribed chart repaints its svg. */
	async function sample(): Promise<void> {
		ensureBarFacePicturePixelSampled(`blob:thumbnail-${sampleIndex++}`, async () => '#123456');
		for (let i = 0; i < 6; i++) {
			await Promise.resolve();
		}
	}
	function row(): HTMLElement {
		const el = document.createElement('div');
		document.body.appendChild(el);
		rows.push(el);
		return el;
	}
	afterEach(() => {
		for (const el of rows.splice(0)) {
			disposePptxRenderStage(el);
			el.remove();
		}
		releasePptxRenderCaches();
	});

	it('owns a chart repaint subscription until disposePptxRenderStage retires it', async () => {
		const kept = row();
		const retired = row();
		kept.append(render([chart]));
		retired.append(render([chart]));
		const keptSvg = kept.querySelector('svg');
		const retiredSvg = retired.querySelector('svg');
		expect(retiredSvg).not.toBeNull();
		// The documented replacement step, while the old box is still in the document, so
		// only the dispose call (not the chart's own disconnected check) can drop it.
		disposePptxRenderStage(retired);
		disposePptxRenderStage(retired);
		await sample();
		expect(kept.querySelector('svg')).not.toBe(keptSvg);
		expect(retired.querySelector('svg')).toBe(retiredSvg);
	});

	it('replaces a thumbnail the documented way without leaving the old chart subscribed', async () => {
		const strip = row();
		strip.append(render([chart]));
		const old = strip.firstElementChild as HTMLElement;
		const oldSvg = old.querySelector('svg');
		disposePptxRenderStage(strip);
		strip.replaceChildren(render([chart]));
		// Park the replaced box in the document: a leaked subscription would repaint it.
		const parked = row();
		parked.append(old);
		const freshSvg = strip.querySelector('svg');
		await sample();
		expect(old.querySelector('svg')).toBe(oldSvg);
		expect(strip.querySelector('svg')).not.toBe(freshSvg);
	});
});
