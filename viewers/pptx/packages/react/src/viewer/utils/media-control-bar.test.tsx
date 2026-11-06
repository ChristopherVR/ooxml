// @vitest-environment happy-dom
/**
 * The authoring canvas follows PowerPoint's Normal view: no transport at rest,
 * a control bar below the shape while it is selected or hovered, and a click on
 * the bar plays the clip without reaching the shape's select/move handlers.
 */
import type { MediaPptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { renderMediaElement } from './media-render';

vi.mock(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

let container: HTMLDivElement;
let root: Root;
let play: ReturnType<typeof vi.spyOn>;

beforeEach(() => {
	play = vi.spyOn(HTMLMediaElement.prototype, 'play').mockResolvedValue();
	vi.spyOn(HTMLMediaElement.prototype, 'pause').mockImplementation(() => {});
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	vi.restoreAllMocks();
});

/** Moves the pointer over (or away from) a 100x100 host box at the origin. */
function pointerOver(host: HTMLElement, inside: boolean): void {
	host.getBoundingClientRect = () => ({ left: 0, right: 100, top: 0, bottom: 100 }) as DOMRect;
	document.dispatchEvent(
		new MouseEvent('pointermove', { clientX: inside ? 50 : 500, clientY: 50 }),
	);
}

const video = {
	id: 'v1',
	type: 'media',
	x: 0,
	y: 0,
	width: 100,
	height: 100,
	mediaType: 'video',
	mediaData: 'data:video/mp4;base64,AAAA',
} as MediaPptxElement;

function mount(controlBar: { selected: boolean } | undefined): void {
	act(() => {
		root.render(
			<div data-pptx-element='true' style={{ position: 'absolute' }}>
				{renderMediaElement(video, new Map(), {
					showTransport: false,
					...(controlBar ? { controlBar } : {}),
				})}
			</div>,
		);
	});
}

const bar = (): HTMLElement | null => container.querySelector('[data-pptx-media-bar]');

describe('media control bar (authoring canvas)', () => {
	it('paints no native transport and no bar when it is not an authoring canvas', () => {
		mount(undefined);
		expect(container.querySelector('video')?.controls).toBe(false);
		expect(bar()).toBeNull();
	});

	it('hides the bar at rest, shows it when selected or hovered, keeps native controls off', () => {
		mount({ selected: false });
		expect(container.querySelector('video')?.controls).toBe(false);
		expect(bar()?.hidden).toBe(true);
		const host = container.querySelector<HTMLElement>('[data-pptx-element]')!;
		act(() => {
			pointerOver(host, true);
		});
		expect(bar()?.hidden).toBe(false);
		act(() => {
			pointerOver(host, false);
		});
		expect(bar()?.hidden).toBe(true);
		mount({ selected: true });
		expect(bar()?.hidden).toBe(false);
	});

	it('plays from the bar without the press reaching the shape', () => {
		mount({ selected: true });
		const host = container.querySelector<HTMLElement>('[data-pptx-element]')!;
		const onDown = vi.fn();
		host.addEventListener('pointerdown', onDown);
		const button = bar()!.querySelector('button')!;
		button.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		button.click();
		expect(onDown).not.toHaveBeenCalled();
		expect(play).toHaveBeenCalled();
	});
});
