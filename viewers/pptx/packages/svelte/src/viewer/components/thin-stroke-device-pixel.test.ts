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
const screen = (px: string) => `max(${px}px, var(--pptx-device-px, 0px))`;

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
		expect(strand?.style.getPropertyValue('stroke-width')).toBe(screen('1'));

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
		expect(style.border).toBe(`${screen('0.5')} solid #000`);
	});
});
