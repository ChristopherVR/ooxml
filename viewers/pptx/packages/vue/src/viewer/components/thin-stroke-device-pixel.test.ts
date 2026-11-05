import { mount } from '@vue/test-utils';
import { screenStrokeWidth } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getShapeFillStrokeStyle } from '../composables/element-style';
import ConnectorRenderer from './ConnectorRenderer.vue';
import ShapeEffectOverlay from './ShapeEffectOverlay.vue';
import SlideStage from './SlideStage.vue';
import TableRenderer from './TableRenderer.vue';

/**
 * Issue #23: a stroke never paints thinner than one device pixel on screen.
 * The stage publishes `--pptx-device-px`; every painted width is
 * `max(<authored>px, var(--pptx-device-px, 0px))`, with the authored width kept
 * as the attribute. Shared helpers decide the value; these pin the Vue wiring.
 */
const HAIRLINE = 0.1 * (96 / 72);
const STEP = 'max(var(--pptx-device-px, 0px), 0.001px)';
const screen = (px: string) => `max(${STEP}, round(${px}px, ${STEP}))`;

function el(overrides: Record<string, unknown>): PptxElement {
	return { id: 'el-1', x: 0, y: 0, width: 400, height: 0, ...overrides } as unknown as PptxElement;
}

describe('vue: thin strokes stay one device pixel on screen', () => {
	it('publishes one device pixel in slide px on the stage', () => {
		const wrapper = mount(SlideStage, {
			props: {
				slide: { id: 's1', elements: [] } as unknown as PptxSlide,
				canvasSize: { width: 960, height: 540 },
				mediaDataUrls: new Map<string, string>(),
				scale: 0.5,
			},
		});
		const stage = wrapper.element as HTMLElement;
		expect(stage.style.getPropertyValue('--pptx-device-px')).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);
	});

	it('paints a hairline `line` preset outline at max(authored, one device pixel)', () => {
		const wrapper = mount(ShapeEffectOverlay, {
			props: {
				element: el({
					type: 'shape',
					shapeType: 'line',
					shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
				}),
			},
		});
		const path = wrapper.get('svg path').element as SVGPathElement;
		expect(path.getAttribute('stroke-width')).toBe(String(HAIRLINE));
		expect(path.style.getPropertyValue('stroke-width')).toBe(screen('0.1333'));
	});

	it('holds a connector strand at one device pixel', () => {
		const wrapper = mount(ConnectorRenderer, {
			props: {
				element: el({
					type: 'connector',
					shapeType: 'straightConnector1',
					height: 40,
					shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
				}),
				zIndex: 0,
			},
		});
		const line = wrapper.get('line').element as SVGLineElement;
		expect(line.getAttribute('stroke-width')).toBe('1');
		expect(line.style.getPropertyValue('stroke-width')).toBe(screenStrokeWidth(HAIRLINE, 1));
	});

	it('holds a table diagonal border at one device pixel', () => {
		const wrapper = mount(TableRenderer, {
			props: {
				element: el({
					type: 'table',
					height: 80,
					tableData: {
						columnWidths: [1],
						rows: [
							{
								cells: [
									{
										text: 'X',
										style: { borderDiagDownColor: '#000000', borderDiagDownWidth: 0.5 },
									},
								],
							},
						],
					},
				}),
				zIndex: 0,
			},
		});
		const line = wrapper.get('svg line').element as SVGLineElement;
		expect(line.style.getPropertyValue('stroke-width')).toBe(screen('0.5'));
	});

	it('emits an inset CSS border at max(authored, one device pixel)', () => {
		const style = getShapeFillStrokeStyle(
			el({
				type: 'shape',
				shapeType: 'rect',
				height: 50,
				shapeStyle: { strokeColor: '#000', strokeWidth: 0.5, lineAlignment: 'in' },
			}),
		);
		expect(style.border).toBe('max(0.5px, var(--pptx-device-px, 0px)) solid #000');
	});

	it('snaps a horizontal connector to the device grid and ends it at its real y', () => {
		const mountLine = (height: number) =>
			mount(ConnectorRenderer, {
				props: {
					element: el({
						type: 'connector',
						shapeType: 'straightConnector1',
						x: 27,
						y: 27,
						height,
						shapeStyle: { strokeColor: '#000000', strokeWidth: 1.3333 },
					}),
					zIndex: 0,
				},
			}).get('line').element as SVGLineElement;
		const flat = mountLine(0);
		expect(flat.getAttribute('y2')).toBe('0');
		expect(flat.style.getPropertyValue('transform')).toContain('round(27px');
		expect(mountLine(40).style.getPropertyValue('transform')).toBe('');
	});
});
