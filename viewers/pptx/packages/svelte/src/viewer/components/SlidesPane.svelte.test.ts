/**
 * `SlidesPane`, the slide thumbnail rail a host composes beside its own
 * `SlideCanvas`, wired through `createSlidesPaneProps` over a real deck
 * loaded by `createViewerState` (see `SlidesPaneHarness.svelte`).
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { SlidesPaneLayoutGalleryRequest } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import { createSlidesPaneProps, SlidesPane } from '../../index';
import type { ViewerStateBag } from '../state/create-viewer-state-types';
import SlidesPaneHarness from './SlidesPaneHarness.svelte';

let fourSlides: Uint8Array;
const cleanups: (() => void)[] = [];

beforeAll(async () => {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 4 });
	fourSlides = await handler.save(data.slides);
	handler.dispose();
});

beforeEach(() => {
	localStorage.clear();
});

afterEach(() => {
	for (const cleanup of cleanups.splice(0).reverse()) {
		cleanup();
	}
	document.body.replaceChildren();
});

function firstSlideNumberDeck(): Uint8Array {
	return new Uint8Array(
		readFileSync(resolve(process.cwd(), '../../../../e2e/pptx/fixtures/first-slide-number.pptx')),
	);
}

interface Mounted {
	state: ViewerStateBag;
	target: HTMLElement;
}

async function mountShell(
	props: {
		source?: Uint8Array;
		editable?: boolean;
		onopenlayoutgallery?: (request: SlidesPaneLayoutGalleryRequest) => void;
		ondirtychange?: (dirty: boolean) => void;
	} = {},
): Promise<Mounted> {
	let captured: ViewerStateBag | undefined;
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(SlidesPaneHarness, {
		target,
		props: {
			source: props.source ?? fourSlides,
			editable: props.editable ?? true,
			onopenlayoutgallery: props.onopenlayoutgallery,
			ondirtychange: props.ondirtychange,
			onready: (state: ViewerStateBag) => {
				captured = state;
			},
		},
	});
	cleanups.push(() => {
		void unmount(instance);
		target.remove();
	});
	const state = captured!;
	await vi.waitFor(
		() => {
			flushSync();
			expect(state.loader.loading).toBeFalsy();
			expect(state.displaySlides.length).toBeGreaterThan(0);
		},
		{ timeout: 15000 },
	);
	flushSync();
	return { state, target };
}

function rows(root: ParentNode): HTMLElement[] {
	return [...root.querySelectorAll<HTMLElement>('[data-pptx-chrome="slide-row"]')];
}

function rightClick(element: Element): void {
	element.dispatchEvent(
		new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 120 }),
	);
	flushSync();
}

/** The open slide menu's rows (the shared context menu renders them in its shadow root). */
function menuItems(): HTMLButtonElement[] {
	return [...document.querySelectorAll('pptx-ui-context-menu')].flatMap((menu) => [
		...(menu.shadowRoot?.querySelectorAll<HTMLButtonElement>('[role="menu"] [role="menuitem"]') ??
			[]),
	]);
}

function menuItem(label: string): HTMLButtonElement | undefined {
	return menuItems().find((item) => item.textContent?.trim() === label);
}

/** Select rows `from`..`to` (click, then Shift-click), as a user does. */
function selectRange(root: ParentNode, from: number, to: number): void {
	rows(root)[from]!.click();
	rows(root)[to]!.dispatchEvent(new MouseEvent('click', { bubbles: true, shiftKey: true }));
	flushSync();
}

describe('slidesPane', () => {
	it('renders a row per slide with the shared DOM markers', async () => {
		const { state, target } = await mountShell();
		const rail = target.querySelector('[data-test-rail]')!;

		expect(rows(rail)).toHaveLength(4);
		for (const row of rows(rail)) {
			expect(row.querySelector('[data-pptx-chrome="slide-frame"]')).not.toBeNull();
		}
		expect(rows(rail)[0]!.getAttribute('aria-current')).toBe('true');
		expect(rail.querySelector('[data-pptx-chrome="slide-footer"] button')).not.toBeNull();
		expect(state.viewer.current).toBe(0);
	});

	it('numbers thumbnails from the deck start number without an enclosing viewer', async () => {
		// `p:presentation/@firstSlideNum="45"`: slide i is numbered 44 + i.
		const { state, target } = await mountShell({ source: firstSlideNumberDeck(), editable: false });
		expect(target.querySelector('[data-test-canvas]')?.textContent).toContain('Slide 45');

		// A separate root: nothing above it provides a field or table context,
		// so the thumbnails rely on the props alone.
		const standalone = document.createElement('div');
		document.body.append(standalone);
		const instance = mount(SlidesPane, { target: standalone, props: createSlidesPaneProps(state) });
		cleanups.push(() => {
			void unmount(instance);
			standalone.remove();
		});
		flushSync();

		const frames = rows(standalone).map((row) =>
			row.querySelector('[data-pptx-chrome="slide-frame"]'),
		);
		expect(frames).toHaveLength(3);
		for (const [at, frame] of frames.entries()) {
			expect(frame?.textContent).toContain(`Slide ${45 + at}`);
		}
		expect(frames[0]?.textContent).not.toContain('Slide 1');
	});

	it('passes the canvas deck-wide field and table contexts to the thumbnails', async () => {
		const { state } = await mountShell({ source: firstSlideNumberDeck(), editable: false });
		const props = createSlidesPaneProps(state);

		expect(props.fieldContext?.firstSlideNumber).toBe(45);
		// The slide number and title are per slide; each thumbnail re-points them.
		// The rail's context leaves those keys present but undefined.
		expect(props.fieldContext).toStrictEqual({
			...state.fieldContext(),
			slideNumber: undefined,
			slideTitle: undefined,
		});
		expect(state.fieldContext().slideNumber).toBe(1);
		expect(props.tableStyleContext?.theme).toBe(state.loader.presentationTheme);
		expect(props.tableStyleContext?.theme?.colorScheme).toBe(state.loader.colorScheme);
		expect(props.tableStyleContext?.tableStyleMap).toBe(state.loader.tableStyleMap);
	});

	it('moves the canvas to a clicked slide without marking the deck dirty', async () => {
		const dirtyChanges: boolean[] = [];
		const { state, target } = await mountShell({
			ondirtychange: (dirty) => dirtyChanges.push(dirty),
		});

		rows(target)[2]!.click();
		flushSync();

		expect(state.viewer.current).toBe(2);
		expect(state.activeSlide?.id).toBe(state.displaySlides[2]!.id);
		expect(rows(target)[2]!.getAttribute('aria-current')).toBe('true');
		expect(rows(target)[0]!.getAttribute('aria-current')).toBeNull();
		expect(state.editor.dirty).toBeFalsy();
		expect(dirtyChanges).not.toContain(true);
	});

	it('makes the slide active and hands the Layout request to the host', async () => {
		const requests: SlidesPaneLayoutGalleryRequest[] = [];
		const { state, target } = await mountShell({
			onopenlayoutgallery: (request) => requests.push(request),
		});

		rightClick(rows(target)[1]!);
		menuItem('Layout')!.click();
		flushSync();

		expect(state.viewer.current).toBe(1);
		expect(requests).toStrictEqual([{ slideIndex: 1, x: 40, y: 120 }]);
		expect(state.editor.dirty).toBeFalsy();
	});

	it('still makes the slide active when the host renders no Layout gallery', async () => {
		const { state, target } = await mountShell();

		rightClick(rows(target)[3]!);
		menuItem('Layout')!.click();
		flushSync();

		expect(state.viewer.current).toBe(3);
	});

	it('duplicates every selected slide, each copy after its source', async () => {
		const { state, target } = await mountShell();
		const ids = state.editor.slides.map((slide) => slide.id);

		selectRange(target, 0, 1);
		rightClick(rows(target)[1]!);
		menuItem('Duplicate 2 Slides')!.click();
		flushSync();

		expect(state.editor.slides).toHaveLength(6);
		const after = state.editor.slides.map((slide) => slide.id);
		expect(after[0]).toBe(ids[0]);
		expect(after[2]).toBe(ids[1]);
		expect(after.slice(4)).toStrictEqual(ids.slice(2));
		expect(rows(target)).toHaveLength(6);
	});

	it('deletes every selected slide', async () => {
		const { state, target } = await mountShell();
		const ids = state.editor.slides.map((slide) => slide.id);

		selectRange(target, 1, 2);
		rightClick(rows(target)[2]!);
		menuItem('Delete 2 Slides')!.click();
		flushSync();

		expect(state.editor.slides.map((slide) => slide.id)).toStrictEqual([ids[0], ids[3]]);
		expect(rows(target)).toHaveLength(2);
		expect(state.editor.dirty).toBeTruthy();
	});

	it('adds a slide from the footer through the editor operations', async () => {
		const { state, target } = await mountShell();

		target.querySelector<HTMLButtonElement>('[data-pptx-chrome="slide-footer"] button')!.click();
		flushSync();

		expect(state.editor.slides).toHaveLength(5);
		expect(rows(target)).toHaveLength(5);
		expect(state.viewer.current).toBe(1);
		expect(state.editor.dirty).toBeTruthy();
	});

	it('is read-only without edit permission', async () => {
		const { target } = await mountShell({ editable: false });

		expect(rows(target)).toHaveLength(4);
		expect(target.querySelector('[data-pptx-chrome="slide-footer"]')).toBeNull();
		rightClick(rows(target)[0]!);
		expect(document.querySelector('pptx-ui-context-menu')).toBeNull();
	});
});
