// @vitest-environment happy-dom
import type { PptxSlide, InkPptxElement } from 'ooxml-core/pptx';
import { collectSlideInk } from 'ooxml-core/pptx/editor/ink';
import { expect, it, vi } from 'vitest';

import { attachDrawCanvas, configurePptxDraw } from '../render/draw-canvas';
import './draw-math';
import { createDrawReplay } from './draw-replay';
import { mountRibbonDrawAdvanced } from './ribbon-draw-advanced';

it('prefills the shared equation fields only on an explicit recognition request', async () => {
	const root = document.createElement('div');
	root.dataset.pptxThemeRoot = '';
	const host = document.createElement('div');
	const viewport = document.createElement('div');
	viewport.dataset.pptxViewport = '';
	const stage = document.createElement('div');
	stage.setAttribute('aria-roledescription', 'slide');
	Object.defineProperties(stage, { offsetWidth: { value: 100 }, offsetHeight: { value: 100 } });
	stage.getBoundingClientRect = () => new DOMRect(0, 0, 100, 100);
	viewport.append(stage);
	root.append(host, viewport);
	document.body.append(root);
	const slide: PptxSlide = {
		id: 'slide',
		rId: 'rId1',
		slideNumber: 1,
		elements: [
			{
				id: 'ink',
				type: 'ink',
				x: 25,
				y: 25,
				width: 20,
				height: 20,
				inkPaths: ['M 0 0 L 20 20'],
			},
		],
	};
	const replaceElements = vi.fn();
	const detach = attachDrawCanvas(root, { getSlide: () => slide, replaceElements });
	const recognizeMath = vi.fn(async () => String.raw`\frac{1}{2}`);
	const mount = mountRibbonDrawAdvanced(
		host,
		() => ({ tool: 'select', width: 2, color: '#000000', editable: true }),
		() => {},
	);
	const math = host.querySelector<HTMLButtonElement>(
		'[data-ribbon-control="draw.convert.inkToMath"]',
	);
	const release = configurePptxDraw(root, { recognizeMath });
	try {
		expect(math?.getAttribute('aria-disabled')).toBe('true');
		host.querySelector<HTMLButtonElement>('[data-ribbon-control="draw.selection.lasso"]')?.click();
		const pointer = (type: string, clientX: number, clientY: number) => {
			stage.dispatchEvent(
				new PointerEvent(type, { bubbles: true, pointerId: 1, clientX, clientY }),
			);
		};
		pointer('pointerdown', 10, 10);
		pointer('pointermove', 55, 10);
		pointer('pointermove', 55, 55);
		pointer('pointermove', 10, 55);
		pointer('pointerup', 10, 10);
		expect(recognizeMath).not.toHaveBeenCalled();
		expect(math?.getAttribute('aria-disabled')).toBe('false');
		math?.click();
		await vi.waitFor(() =>
			expect(root.querySelector('office-ui-dialog textarea')?.value).toBe(String.raw`\frac{1}{2}`),
		);
		expect(recognizeMath).toHaveBeenCalledExactlyOnceWith(
			[
				[
					{ x: 25, y: 25 },
					{ x: 45, y: 45 },
				],
			],
			expect.any(AbortSignal),
		);
		const dialog = root.querySelector('office-ui-dialog');
		expect(dialog?.hasAttribute('open')).toBe(true);
		const footer = dialog?.querySelector('office-ui-dialog-footer');
		const insert = [...(footer?.shadowRoot?.querySelectorAll('button') ?? [])].find(
			(button) => button.textContent?.trim() === 'Insert',
		);
		expect(insert).toBeDefined();
		insert?.click();
		expect(replaceElements).toHaveBeenCalledExactlyOnceWith(
			slide,
			[
				expect.objectContaining({
					type: 'shape',
					textSegments: [expect.objectContaining({ equationXml: expect.any(Object) })],
				}),
			],
			expect.any(String),
		);
	} finally {
		mount.dispose();
		release();
		detach();
		root.remove();
	}
});

it('seeks backward across strokes and restores the canvas when replay closes', () => {
	const stage = document.createElement('div');
	const original = document.createElement('div');
	original.dataset.elementId = 'ink';
	stage.append(original);
	const owner: PptxSlide['elements'][number] = {
		id: 'ink',
		type: 'ink',
		x: 0,
		y: 0,
		width: 100,
		height: 20,
		inkPaths: ['M 0 0 L 100 0', 'M 0 0 L 100 0'],
	};
	const replay = createDrawReplay(
		stage,
		[100, 500].map((start, index) => ({
			owner,
			index,
			points: [
				{ x: 0, y: 0 },
				{ x: 80, y: 0 },
				{ x: 100, y: 0 },
			],
			style: {
				color: '#000000',
				width: 2,
				opacity: 1,
				pointTimestamps: [start, start + 10, start + 100],
			},
		})),
		() => {},
	);
	try {
		const paths = [...stage.querySelectorAll('path')];
		expect(original.style.visibility).toBe('hidden');
		replay?.seek(500);
		expect(paths.map((path) => path.getAttribute('stroke-dashoffset'))).toStrictEqual(['0', '0']);
		replay?.seek(10);
		expect(Number(paths[0]?.getAttribute('stroke-dashoffset'))).toBeCloseTo(0.2);
		expect(paths[1]?.getAttribute('visibility')).toBe('hidden');
	} finally {
		replay?.dispose();
	}
	expect(stage.querySelector('svg')).toBeNull();
	expect(original.style.visibility).toBe('');
});

it('leaves an owner visible when the collection budget cuts through its strokes', () => {
	const stage = document.createElement('div');
	const firstNode = document.createElement('div');
	firstNode.dataset.elementId = 'complete';
	const secondNode = document.createElement('div');
	secondNode.dataset.elementId = 'partial';
	stage.append(firstNode, secondNode);
	const first: InkPptxElement = {
		id: 'complete',
		type: 'ink',
		x: 0,
		y: 0,
		width: 20,
		height: 20,
		inkPaths: ['M 0 0 L 20 20'],
	};
	const second: InkPptxElement = {
		...first,
		id: 'partial',
		inkPaths: Array.from({ length: 1000 }, () => 'M 0 0 L 20 20'),
	};
	const collected = collectSlideInk([first, second]);
	expect(collected).toHaveLength(1000);
	const replay = createDrawReplay(stage, collected, () => {});
	try {
		expect(firstNode.style.visibility).toBe('hidden');
		expect(secondNode.style.visibility).toBe('');
		expect(stage.querySelectorAll('[data-pptx-draw-replay] path')).toHaveLength(1);
	} finally {
		replay?.dispose();
	}
	expect(firstNode.style.visibility).toBe('');
});
