// @vitest-environment jsdom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTranslator } from './i18n';
import { createDefaultRegistry, renderSlideStage } from './render';

/**
 * Issue #23: a stroke never paints thinner than one device pixel on screen.
 * This binding paints through the shared DOM renderer, so this pins that the
 * stage it renders (editor, show and thumbnails alike) publishes the device
 * pixel and that a hairline is painted against it, authored width kept as the
 * attribute for export.
 */
const HAIRLINE = 0.1 * (96 / 72);

describe('vanilla: thin strokes stay one device pixel on screen', () => {
	it('publishes the device pixel on the stage and holds a hairline to it', () => {
		const line = {
			type: 'shape',
			id: 'line',
			x: 10,
			y: 10,
			width: 400,
			height: 0,
			shapeType: 'line',
			shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
		} as PptxElement;
		const stage = renderSlideStage({
			document,
			slide: { id: 's1', rId: 'r1', slideNumber: 1, elements: [line] } as PptxSlide,
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map(),
			registry: createDefaultRegistry(),
			t: createTranslator('en'),
			scale: 0.25,
			interactive: true,
		});
		expect(stage.style.getPropertyValue('--pptx-device-px')).toBe(
			'calc(1px / (0.25 * var(--pptx-dpr, 1)))',
		);
		const path = stage.querySelector<SVGPathElement>('[data-element-id="line"] svg path');
		expect(path?.getAttribute('stroke-width')).toBe(String(HAIRLINE));
		expect(path?.style.getPropertyValue('stroke-width')).toBe(
			'max(0.1333px, var(--pptx-device-px, 0px))',
		);
	});
});
