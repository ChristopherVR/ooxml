// @vitest-environment happy-dom
import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../../i18n/translator';
import { createElementRendererRegistry } from '../registry';
import type { ElementRenderContext } from '../types';
import { renderMediaElement } from './media';

/**
 * The vanilla stage follows PowerPoint's Normal view: no transport at rest and a
 * control bar under a selected or hovered video, none on a show, a still of a
 * slide or the reading view (which keeps the native transport).
 */
function context(overrides: Partial<ElementRenderContext> = {}): ElementRenderContext {
	const registry = createElementRendererRegistry();
	const base = {
		document,
		slide: { id: 's1', rId: 'rId1', slideNumber: 1, elements: [] },
		canvasSize: { width: 1280, height: 720 },
		scale: 1,
		mediaDataUrls: new Map<string, string>(),
		t: createTranslator(),
		smartArt3D: false,
		surfaceChart3D: false,
		barChart3D: false,
		lineChart3D: false,
		areaChart3D: false,
		pieChart3D: false,
		presenting: false,
		interactive: true,
		registry,
		...overrides,
	} as ElementRenderContext;
	return { ...base, renderElement: (el, z) => registry.resolve(el.type)(el, z, base) };
}

/** Moves the pointer over (or away from) a 100x100 host box at the origin. */
function pointerOver(host: HTMLElement, inside: boolean): void {
	host.getBoundingClientRect = () => ({ left: 0, right: 100, top: 0, bottom: 100 }) as DOMRect;
	document.dispatchEvent(
		new MouseEvent('pointermove', { clientX: inside ? 50 : 500, clientY: 50 }),
	);
}

const video = {
	type: 'media',
	id: 'v1',
	x: 0,
	y: 0,
	width: 100,
	height: 100,
	mediaType: 'video',
	mediaData: 'data:video/mp4;base64,AAAA',
} as PptxElement;

const bar = (node: HTMLElement) => node.querySelector<HTMLElement>('[data-pptx-media-bar]');

describe('vanilla media control bar', () => {
	it('is hidden at rest, with no native transport', () => {
		const node = renderMediaElement(video, 0, context()) as HTMLElement;
		expect(bar(node)?.hidden).toBe(true);
		expect(node.querySelector('video')?.controls).toBe(false);
	});

	it('shows for a selected element and on hover', () => {
		const selected = renderMediaElement(
			video,
			0,
			context({ selectedElementIds: new Set(['v1']) }),
		) as HTMLElement;
		expect(bar(selected)?.hidden).toBe(false);
		const node = renderMediaElement(video, 0, context()) as HTMLElement;
		pointerOver(node, true);
		expect(bar(node)?.hidden).toBe(false);
	});

	it('is absent on a show, on a still and in the reading view', () => {
		for (const overrides of [
			{ presenting: true, interactive: false },
			{ presenting: false, interactive: false },
			{ presenting: false, interactive: false, reading: true },
		]) {
			const node = renderMediaElement(video, 0, context(overrides)) as HTMLElement;
			expect(bar(node)).toBeNull();
		}
	});
});
