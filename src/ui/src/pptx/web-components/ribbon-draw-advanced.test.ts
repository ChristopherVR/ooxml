// @vitest-environment happy-dom
import type { PptxSlide } from 'ooxml-core/pptx';
import { expect, it, vi } from 'vitest';

import { attachDrawCanvas, configurePptxDraw } from '../render/draw-canvas';
import './draw-math';
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
			expect(root.querySelector('dialog textarea')?.value).toBe(String.raw`\frac{1}{2}`),
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
		const insert = [...root.querySelectorAll<HTMLButtonElement>('dialog button')].find(
			(button) => button.textContent === 'Insert',
		);
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
