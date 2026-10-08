// @vitest-environment jsdom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n/translator';
import { createDefaultRegistry } from './elements';
import { renderSlideStage } from './slide-stage';

/**
 * Issue #23: a 0.1 pt line must not fade away on screen. The stage publishes
 * one device pixel in slide px and every painted stroke / border is written as
 * `max(<authored>px, var(--pptx-device-px, 0px))`, so outside a stage (export,
 * print) the authored width is kept.
 */
const HAIRLINE = 0.1 * (96 / 72);

function el(overrides: Partial<PptxElement> & Pick<PptxElement, 'type' | 'id'>): PptxElement {
	return { x: 10, y: 10, width: 400, height: 0, ...overrides } as PptxElement;
}

function stage(elements: PptxElement[], scale = 0.5): HTMLElement {
	const slide: PptxSlide = { id: 's1', rId: 'r1', slideNumber: 1, elements };
	return renderSlideStage({
		document,
		slide,
		canvasSize: { width: 1280, height: 720 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator(),
		scale,
	});
}

const screenWidth = (authored: string): string => `max(${authored}px, var(--pptx-device-px, 0px))`;

describe('one device pixel minimum (vanilla DOM renderer)', () => {
	it('publishes the device pixel on the stage for its scale', () => {
		const node = stage([], 0.5);
		expect(node.style.getPropertyValue('--pptx-device-px')).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);
	});

	it('keeps a hairline `line` preset at one device pixel, authored width as the attribute', () => {
		const node = stage([
			el({
				type: 'shape',
				id: 'line',
				shapeType: 'line',
				shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
			}),
		]);
		const path = node.querySelector<SVGPathElement>('[data-element-id="line"] svg path');
		expect(path?.getAttribute('stroke-width')).toBe(String(HAIRLINE));
		expect(path?.style.getPropertyValue('stroke-width')).toBe(screenWidth('0.1333'));
	});

	it('paints a no-line shape with no stroke at all', () => {
		const node = stage([
			el({
				type: 'shape',
				id: 'box',
				shapeType: 'rect',
				height: 50,
				shapeStyle: { fillColor: '#ff0000' },
			}),
		]);
		const box = node.querySelector<HTMLElement>('[data-element-id="box"]');
		expect(box?.querySelector('svg path[stroke]')).toBeNull();
		expect(box?.style.borderTopWidth || '0px').toBe('0px');
	});

	it('holds an inset (`algn="in"`) CSS border at one device pixel', () => {
		const node = stage([
			el({
				type: 'shape',
				id: 'inset',
				shapeType: 'rect',
				height: 50,
				shapeStyle: { strokeColor: '#000000', strokeWidth: 0.5, lineAlignment: 'in' },
			}),
		]);
		const box = node.querySelector<HTMLElement>('[data-element-id="inset"]');
		expect(box?.getAttribute('style')).toContain(screenWidth('0.5'));
	});

	it('holds a connector strand at one device pixel, never under its 1px floor', () => {
		const node = stage([
			el({
				type: 'connector',
				id: 'conn',
				height: 60,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
			}),
		]);
		const line = node.querySelector<SVGLineElement>('[data-element-id="conn"] svg line');
		expect(line?.getAttribute('stroke-width')).toBe('1');
		expect(line?.style.getPropertyValue('stroke-width')).toBe(screenWidth('1'));
	});

	it('holds table cell borders and diagonals at one device pixel', () => {
		const node = stage([
			el({
				type: 'table',
				id: 'tbl',
				height: 80,
				tableData: {
					columnWidths: [1],
					rows: [
						{
							height: 80,
							cells: [
								{
									text: 'x',
									style: {
										borderTopWidth: 0.5,
										borderTopColor: '#000000',
										borderDiagDownColor: '#000000',
										borderDiagDownWidth: 0.5,
									},
								},
							],
						},
					],
				},
			} as Partial<PptxElement> & Pick<PptxElement, 'type' | 'id'>),
		]);
		const td = node.querySelector<HTMLElement>('[data-element-id="tbl"] td');
		expect(td?.getAttribute('style')).toContain(`border-top: ${screenWidth('0.5')} solid`);
		const diagonal = td?.querySelector<SVGLineElement>('svg line');
		expect(diagonal?.getAttribute('stroke-width')).toBe('0.5');
		expect(diagonal?.style.getPropertyValue('stroke-width')).toBe(screenWidth('0.5'));
	});
});
