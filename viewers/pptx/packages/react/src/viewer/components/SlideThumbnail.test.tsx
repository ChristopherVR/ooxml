// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

// Import through the package entry so the test also covers the public export.
import { PptxI18nProvider, SlideThumbnail } from '../../index';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

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
const stage = () =>
	container.firstElementChild?.firstElementChild as HTMLElement | null | undefined;
// Thumbnails draw elements without their ids, so ids stay unique to the editing canvas.
const drawn = () => [
	...(stage()?.querySelectorAll<HTMLElement>(':scope > [data-static-element-type]') ?? []),
];

describe('slideThumbnail', () => {
	it('draws the slide stage scaled to the width the host picks, at the deck aspect ratio', () => {
		act(() =>
			root.render(
				<SlideThumbnail
					slide={slide([shape('a')])}
					templateElements={[]}
					canvasSize={canvasSize}
					previewWidth={120}
				/>,
			),
		);
		const box = container.firstElementChild as HTMLElement;
		expect(box.style.height).toBe('67.5px');
		expect(stage()?.style.width).toBe('960px');
		expect(stage()?.style.transform).toBe('scale(0.125)');
	});

	it('draws inherited template elements beneath the slide elements', () => {
		act(() =>
			root.render(
				<SlideThumbnail
					slide={slide([shape('own', 200)])}
					templateElements={[shape('layout', 10)]}
					canvasSize={canvasSize}
					previewWidth={120}
				/>,
			),
		);
		expect(drawn().map((node) => node.style.left)).toStrictEqual(['10px', '200px']);
	});

	it('follows edits to the slide model without a save', () => {
		const render = (elements: PptxElement[]) =>
			act(() =>
				root.render(
					<SlideThumbnail
						slide={slide(elements)}
						templateElements={[]}
						canvasSize={canvasSize}
						previewWidth={120}
					/>,
				),
			);
		render([shape('a')]);
		expect(drawn()).toHaveLength(1);
		render([shape('a'), shape('b')]);
		expect(drawn()).toHaveLength(2);
	});

	it('follows slide metadata edits that keep the same elements', () => {
		const number = {
			...shape('number'),
			type: 'text',
			textSegments: [{ text: '#', style: {}, fieldType: 'slidenum' }],
		} as unknown as PptxElement;
		const base = slide([number]);
		// Stable context and template references, as a host keeps them: only the
		// slide object changes below.
		const fieldContext = {};
		const templateElements: PptxElement[] = [];
		const render = (next: PptxSlide) =>
			act(() =>
				root.render(
					<SlideThumbnail
						slide={next}
						templateElements={templateElements}
						canvasSize={canvasSize}
						previewWidth={120}
						fieldContext={fieldContext}
					/>,
				),
			);
		render(base);
		expect(container.textContent).toBe('1');
		// Adding a fade keeps the element and template references: the badge must still show.
		render({ ...base, transition: { type: 'fade' } } as PptxSlide);
		expect(container.querySelector('[title]')?.textContent).toBeTruthy();
		// Renumbering (a slide moved earlier) keeps them too: the number field must follow.
		render({ ...base, slideNumber: 7 });
		expect(container.textContent).toBe('7');
	});

	// The same badge, from the same shared descriptor, in every binding.
	it('draws the transition badge in the corner, at the slide scale', () => {
		const translate = (key: string, params?: Record<string, unknown>) =>
			({
				'pptx.ribbon.transition.fade': 'Fade',
				'pptx.slideThumbnail.transitionTooltip': `Transition: ${String(params?.name)}`,
			})[key] ?? key;
		const render = (transition: PptxSlide['transition']) =>
			act(() =>
				root.render(
					<PptxI18nProvider translate={translate} locale='en'>
						<SlideThumbnail
							slide={{ ...slide([shape('a')]), transition } as PptxSlide}
							templateElements={[]}
							canvasSize={canvasSize}
							previewWidth={120}
						/>
					</PptxI18nProvider>,
				),
			);
		const badge = () =>
			container.querySelector<HTMLElement>(':scope > div > [data-pptx-thumbnail-transition]');
		render({ type: 'fade' });
		expect(badge()?.textContent).toBe('Fade');
		expect(badge()?.title).toBe('Transition: Fade');
		expect(badge()?.style.transform).toBe('scale(0.125)');
		// The drawn slide's corner, 120px from the box's left edge.
		expect(badge()?.style.right).toBe('calc(100% - 119.75px)');
		render({ type: 'cut' });
		expect(badge()).toBeNull();
		render(undefined);
		expect(badge()).toBeNull();
	});
});
