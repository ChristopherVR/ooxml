// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n/translator';
import { createDefaultRegistry } from './elements';
import { renderSlideStage } from './slide-stage';
import type { SlideStageOptions } from './slide-stage';

const ORIGIN = 'https://linked-content.test';

/** A slide whose background, picture, picture fill and table cell fill are linked to {@link ORIGIN}. */
function linkedSlide(): PptxSlide {
	return {
		id: 'slide-1',
		rId: 'rId1',
		slideNumber: 1,
		backgroundImage: `${ORIGIN}/background.png`,
		elements: [
			{
				type: 'picture',
				id: 'picture',
				x: 0,
				y: 0,
				width: 100,
				height: 100,
				imagePath: `${ORIGIN}/picture.png`,
				imageData: `${ORIGIN}/picture.png`,
			},
			{
				type: 'shape',
				id: 'fill',
				x: 120,
				y: 0,
				width: 100,
				height: 100,
				shapeType: 'rect',
				shapeStyle: { fillMode: 'image', fillImageUrl: `${ORIGIN}/fill.png` },
			},
		],
	} as PptxSlide;
}

function stage(options: Partial<SlideStageOptions> = {}): HTMLElement {
	return renderSlideStage({
		document,
		slide: linkedSlide(),
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator(),
		...options,
	});
}

/** Whether any attribute of the stage would make the browser contact {@link ORIGIN}. */
const contacts = (node: HTMLElement): boolean =>
	[node, ...node.querySelectorAll('*')].some((el) =>
		[...el.attributes].some((attribute) => attribute.value.includes(ORIGIN)),
	);

describe('renderSlideStage linked content', () => {
	it('draws nothing linked unless the viewer allows it', () => {
		expect(contacts(stage())).toBe(false);
		expect(contacts(stage({ linkedContent: false }))).toBe(false);
	});

	it('draws the linked background, picture and fill once allowed', () => {
		const node = stage({ linkedContent: true });
		expect(node.outerHTML).toContain(`${ORIGIN}/background.png`);
		expect(node.outerHTML).toContain(`${ORIGIN}/picture.png`);
		expect(node.outerHTML).toContain(`${ORIGIN}/fill.png`);
	});

	it('hands edit callbacks the model element, links included', () => {
		const slide = linkedSlide();
		const onTableResizeColumns = vi.fn<(element: PptxElement, widths: number[]) => void>();
		renderSlideStage({
			document,
			slide,
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map(),
			registry: {
				...createDefaultRegistry(),
				resolve: () => (element, _z, context) => {
					context.onTableResizeColumns?.(element, [1]);
					return document.createElement('div');
				},
			},
			t: createTranslator(),
			onTableResizeColumns,
		});
		expect(onTableResizeColumns.mock.calls.map(([element]) => element)).toStrictEqual(
			slide.elements,
		);
		expect(onTableResizeColumns.mock.calls[0]![0]).toBe(slide.elements[0]);
	});
});
