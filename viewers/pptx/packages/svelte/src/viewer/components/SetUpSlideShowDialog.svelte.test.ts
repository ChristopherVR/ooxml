import type { PptxCustomShow, PptxPresentationProperties } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { allButtons } from './dialog-footer.test-support';
import SetUpSlideShowDialog from './SetUpSlideShowDialog.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const SHOWS: PptxCustomShow[] = [{ id: '4', name: 'Short', slideRIds: ['rId1'] }];

function saved(act: (radios: HTMLElement[]) => void): PptxPresentationProperties {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const onsave = vi.fn();
	const instance = mount(SetUpSlideShowDialog, {
		target,
		props: { properties: {}, customShows: SHOWS, slideCount: 6, onclose: () => {}, onsave },
	});
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	flushSync();
	// Order: show type x3, all slides, from/to, custom show, advance x2.
	act(Array.from(target.querySelectorAll<HTMLElement>('pptx-ui-radio')));
	flushSync();
	allButtons(target)
		.find((button) => button.textContent?.trim() === 'OK')
		?.click();
	expect(onsave).toHaveBeenCalledTimes(1);
	return onsave.mock.calls[0]?.[0] as PptxPresentationProperties;
}

const change = (radio: HTMLElement | undefined) => {
	radio?.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('svelte set up slide show dialog follows the shared rules', () => {
	it('a kiosk show loops continuously', () => {
		expect(saved((radios) => change(radios[2]))).toMatchObject({
			showType: 'kiosk',
			loopContinuously: true,
		});
	});

	it('a range starts as the whole deck', () => {
		expect(saved((radios) => change(radios[4]))).toMatchObject({
			showSlidesMode: 'range',
			showSlidesFrom: 1,
			showSlidesTo: 6,
		});
	});

	it('a custom show defaults to the first show', () => {
		expect(saved((radios) => change(radios[5]))).toMatchObject({
			showSlidesMode: 'customShow',
			showSlidesCustomShowId: '4',
		});
	});
});
