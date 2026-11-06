// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	attachMediaControlBar,
	mediaControlBarEnabled,
	mediaControlBarVisible,
} from './media-control-bar';

afterEach(() => {
	vi.restoreAllMocks();
});

/** Moves the pointer over (or away from) a 100x100 host box at the origin. */
function pointerOver(host: HTMLElement, inside: boolean): void {
	host.getBoundingClientRect = () => ({ left: 0, right: 100, top: 0, bottom: 100 }) as DOMRect;
	document.dispatchEvent(
		new MouseEvent('pointermove', { clientX: inside ? 50 : 500, clientY: 50 }),
	);
}

describe('mediaControlBarEnabled', () => {
	it('is offered on the editable canvas only', () => {
		expect(mediaControlBarEnabled({ presenting: false, preview: false })).toBe(true);
		expect(mediaControlBarEnabled({ presenting: true, preview: false })).toBe(false);
		expect(mediaControlBarEnabled({ presenting: false, preview: true })).toBe(false);
		expect(mediaControlBarEnabled({ presenting: false, preview: false, reading: true })).toBe(
			false,
		);
	});
});

describe('mediaControlBarVisible', () => {
	it('is never visible at rest, visible when selected or hovered', () => {
		expect(mediaControlBarVisible({ selected: false, hovered: false })).toBe(false);
		expect(mediaControlBarVisible({ selected: true, hovered: false })).toBe(true);
		expect(mediaControlBarVisible({ selected: false, hovered: true })).toBe(true);
	});
});

describe('attachMediaControlBar', () => {
	function setup(selected = false) {
		const host = document.createElement('div');
		const media = document.createElement('video');
		host.append(media);
		document.body.append(host);
		const bar = attachMediaControlBar(host, media, { t: (key) => key, selected });
		return { host, media, bar };
	}

	it('hides at rest, follows hover and selection, and removes itself', () => {
		const { host, bar } = setup();
		expect(bar.element.hidden).toBe(true);
		pointerOver(host, true);
		expect(bar.element.hidden).toBe(false);
		pointerOver(host, false);
		expect(bar.element.hidden).toBe(true);
		bar.setSelected(true);
		expect(bar.element.hidden).toBe(false);
		bar.dispose();
		expect(host.querySelector('[data-pptx-media-bar]')).toBeNull();
		host.remove();
	});

	it('plays and pauses from the bar and keeps the press away from the shape', () => {
		const { host, media, bar } = setup(true);
		const play = vi.spyOn(media, 'play').mockResolvedValue();
		const onDown = vi.fn();
		host.addEventListener('pointerdown', onDown);
		const button = bar.element.querySelector('button')!;
		button.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		button.click();
		expect(onDown).not.toHaveBeenCalled();
		expect(play).toHaveBeenCalledTimes(1);
		bar.dispose();
		host.remove();
	});

	it('is chrome: marked so exports and stills can strip it', () => {
		const { host, bar } = setup(true);
		expect(bar.element.getAttribute('data-pptx-media-chrome')).toBe('control-bar');
		bar.dispose();
		host.remove();
	});
});
