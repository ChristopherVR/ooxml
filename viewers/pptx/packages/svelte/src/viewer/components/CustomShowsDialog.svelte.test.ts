/**
 * The Custom Shows dialog's active-show picker.
 *
 * The dialog could always DEFINE shows; what it could not do was select one, so
 * a custom show had no effect on what presented. The picker carries React's
 * labels because it is the same control, hosted in this dialog rather than in
 * the ribbon's primary row.
 */
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import CustomShowsDialog from './CustomShowsDialog.svelte';
import { allButtons } from './dialog-footer.test-support';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const SHOWS: PptxCustomShow[] = [
	{ id: 'show-1', name: 'Short', slideRIds: ['rId1', 'rId3'] },
	{ id: 'show-2', name: 'Investors', slideRIds: ['rId2'] },
];

const SLIDES = Array.from(
	{ length: 3 },
	(_unused, index) =>
		({
			id: `s${index + 1}`,
			rId: `rId${index + 1}`,
			slideNumber: index + 1,
			elements: [],
		}) as PptxSlide,
);

function open(activeShowId: string | null) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const onsetactive = vi.fn();
	const onsave = vi.fn();
	const instance = mount(CustomShowsDialog, {
		target,
		props: {
			shows: SHOWS,
			slides: SLIDES,
			activeShowId,
			onclose: () => undefined,
			onsave,
			onsetactive,
		},
	});
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	flushSync();
	const select = target.querySelector<HTMLSelectElement>('.active-picker pptx-ui-select')!;
	return { target, select, onsetactive, onsave };
}

describe('customShowsDialog active-show picker', () => {
	it('offers All Slides plus every defined show, at React parity labels', () => {
		const { select } = open(null);
		expect(select.getAttribute('aria-label')).toBe('Select custom show');
		expect([...select.options].map((option) => option.textContent)).toStrictEqual([
			'All Slides',
			'Short',
			'Investors',
		]);
	});

	it('reflects the show already selected', () => {
		expect(open('show-2').select.value).toBe('show-2');
	});

	it('selecting a show restricts playback; All Slides lifts the restriction', () => {
		const { select, onsetactive } = open(null);
		select.value = 'show-1';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();
		expect(onsetactive).toHaveBeenLastCalledWith('show-1');
		select.value = '';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();
		expect(onsetactive).toHaveBeenLastCalledWith(null);
	});

	it('re-points the restriction on save so a deleted show cannot stay selected', () => {
		const { target, select, onsetactive } = open('show-1');
		select.value = 'show-1';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();
		allButtons(target)
			.find((button) => button.textContent === 'Remove')!
			.click();
		flushSync();
		allButtons(target)
			.find((button) => button.textContent === 'OK')!
			.click();
		flushSync();
		expect(onsetactive).toHaveBeenLastCalledWith(null);
	});

	it('copies a show, moves its slides and clears jumps into a removed show', () => {
		const { target, onsave } = open('show-1');
		const click = (label: string) => {
			allButtons(target)
				.find((button) => button.textContent?.trim() === label)!
				.click();
			flushSync();
		};
		click('Copy');
		const rows = () => [
			...target.querySelectorAll<HTMLElement>('[aria-label="Slides in show"] li'),
		];
		rows()[1]!.click();
		flushSync();
		click('Move up');
		click('OK');
		const [shows] = onsave.mock.calls[0] as [PptxCustomShow[], unknown];
		expect(shows.map((show) => show.name)).toStrictEqual(['Short', 'Investors', 'Copy of Short']);
		expect(shows[2]!.slideRIds).toStrictEqual(['rId3', 'rId1']);
	});
});
