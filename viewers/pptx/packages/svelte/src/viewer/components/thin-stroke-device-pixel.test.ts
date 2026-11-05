import { screenStrokeWidth } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { getShapeFillStrokeStyle } from '../style/element-style';
import SlideStage from './SlideStage.svelte';

/**
 * Issue #23: a stroke never paints thinner than one device pixel on screen.
 * The stage publishes `--pptx-device-px`; every painted width is
 * `max(<authored>px, var(--pptx-device-px, 0px))`, with the authored width kept
 * as the attribute. Shared helpers decide the value; these pin the Svelte wiring.
 */
const HAIRLINE = 0.1 * (96 / 72);
const STEP = 'max(var(--pptx-device-px, 0px), 0.001px)';
const screen = (px: string) => `max(${STEP}, round(${px}px, ${STEP}))`;

function el(overrides: Record<string, unknown>): PptxElement {
	return { x: 0, y: 0, width: 400, height: 0, ...overrides } as unknown as PptxElement;
}

let cleanup: (() => void) | undefined;

function mountStage(elements: PptxElement[]): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(SlideStage, {
		target,
		props: {
			slide: { id: 's1', rId: 'r1', slideNumber: 1, elements } as PptxSlide,
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map<string, string>(),
			scale: 0.5,
			interactive: true,
		},
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

describe('svelte: thin strokes stay one device pixel on screen', () => {
	it('publishes the device pixel on the stage and holds every stroke kind to it', () => {
		const root = mountStage([
			el({
				type: 'shape',
				id: 'line',
				shapeType: 'line',
				shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
			}),
			el({
				type: 'connector',
				id: 'conn',
				y: 100,
				height: 40,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
			}),
			el({
				type: 'table',
				id: 'tbl',
				y: 200,
				height: 80,
				tableData: {
					columnWidths: [1],
					rows: [
						{
							cells: [
								{ text: 'X', style: { borderDiagDownColor: '#000000', borderDiagDownWidth: 0.5 } },
							],
						},
					],
				},
			}),
		]);
		const stage = root.querySelector<HTMLElement>('.pptx-svelte-stage');
		expect(stage?.style.getPropertyValue('--pptx-device-px')).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);

		const outline = root.querySelector<SVGPathElement>('[data-element-id="line"] svg path[stroke]');
		expect(outline?.getAttribute('stroke-width')).toBe(String(HAIRLINE));
		expect(outline?.style.getPropertyValue('stroke-width')).toBe(screen('0.1333'));

		const strand = root.querySelector<SVGLineElement>('[data-element-id="conn"] svg line');
		expect(strand?.getAttribute('stroke-width')).toBe('1');
		expect(strand?.style.getPropertyValue('stroke-width')).toBe(screenStrokeWidth(HAIRLINE, 1));

		const diagonal = root.querySelector<SVGLineElement>('[data-element-id="tbl"] td svg line');
		expect(diagonal?.style.getPropertyValue('stroke-width')).toBe(screen('0.5'));
	});

	it('emits an inset CSS border at max(authored, one device pixel)', () => {
		const style = getShapeFillStrokeStyle(
			el({
				type: 'shape',
				id: 'inset',
				shapeType: 'rect',
				height: 50,
				shapeStyle: { strokeColor: '#000', strokeWidth: 0.5, lineAlignment: 'in' },
			}),
		);
		expect(style.border).toBe('max(0.5px, var(--pptx-device-px, 0px)) solid #000');
	});

	it('snaps a horizontal connector to the device grid and ends it at its real y', () => {
		const root = mountStage([
			el({
				type: 'connector',
				id: 'flat',
				x: 27,
				y: 27,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: 1.3333 },
			}),
			el({
				type: 'connector',
				id: 'diag',
				y: 100,
				height: 40,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: 1.3333 },
			}),
		]);
		const flat = root.querySelector<SVGElement>('[data-element-id="flat"] svg line');
		expect(flat?.getAttribute('y2')).toBe('0');
		expect(flat?.style.getPropertyValue('transform')).toContain('round(27px');
		const diag = root.querySelector<SVGElement>('[data-element-id="diag"] svg line');
		expect(diag?.style.getPropertyValue('transform')).toBe('');
	});

	it('snaps every edge of a plain rectangle outline, from its exact EMU position', () => {
		const root = mountStage([
			el({
				type: 'shape',
				id: 'rect',
				shapeType: 'rect',
				x: 454,
				y: 387,
				width: 80,
				height: 67,
				xEmu: 4321810,
				yEmu: 3688080,
				widthEmu: 762000,
				heightEmu: 641350,
				shapeStyle: { strokeColor: '#000000', strokeWidth: 1 },
			}),
		]);
		const path = root.querySelector<SVGElement>('[data-element-id="rect"] svg[data-pptx-outline] path');
		const transform = path?.getAttribute('style') ?? '';
		expect(transform).toContain('scale(calc((80px');
		expect(transform).toContain('(67.3333px');
	});

	it('ends a connector with a flat cap unless the line declares another', () => {
		const root = mountStage([
			el({
				type: 'connector',
				id: 'plain',
				x: 27,
				y: 27,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: 12 },
			}),
			el({
				type: 'connector',
				id: 'round',
				x: 27,
				y: 127,
				shapeType: 'straightConnector1',
				shapeStyle: { strokeColor: '#000000', strokeWidth: 12, lineCap: 'rnd' },
			}),
		]);
		const cap = (id: string) =>
			root
				.querySelector<SVGElement>(`[data-element-id="${id}"] svg line[data-pptx-line]`)
				?.getAttribute('stroke-linecap');
		expect(cap('plain')).toBe('butt');
		expect(cap('round')).toBe('round');
	});
});
