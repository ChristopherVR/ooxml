import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { CanvasSize } from 'ooxml-ui/pptx';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import { allButtons } from '../../dialog-footer.test-support';
import InsertTab from './InsertTab.svelte';

/**
 * InsertTab tests: the adapter wiring over the shared `pptx-ui-ribbon-insert`.
 * The shared element has its own contract tests; these pin that each typed
 * intent reaches the native editor mutation or dialog, and that read-only and
 * selection gating are reflected.
 */

registerPptxWebControls();

const CANVAS: CanvasSize = { width: 960, height: 540 };

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function makeEditor(editable = true): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = editable;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [] }]);
	return editor;
}

function mountTab(editor: EditorState, extra: Record<string, unknown> = {}): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(InsertTab, { target, props: { editor, canvasSize: CANVAS, ...extra } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

function control(target: HTMLElement, id: string): HTMLButtonElement {
	return target
		.querySelector(`[data-ribbon-control="${id}"]`)!
		.shadowRoot!.querySelector('button')!;
}

function headerFooter(target: HTMLElement): HTMLElement {
	return target.querySelector<HTMLElement>('[data-ribbon-control="insert.text.headerFooter"]')!;
}

function pickItem(target: HTMLElement, id: string, item: string): void {
	target.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!.click();
	target
		.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] [data-insert-item="${item}"]`)!
		.click();
}

function fireChange(input: HTMLInputElement, file: File): void {
	Object.defineProperty(input, 'files', { value: [file], configurable: true });
	input.dispatchEvent(new Event('change', { bubbles: true }));
}

describe('insertTab', () => {
	it('inserts a text box and a 3x3 table from the shared commands', () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		control(target, 'insert.text.textBox').click();
		pickItem(target, 'insert.tables.table', 'insertTable');
		flushSync();
		expect(editor.slides[0]?.elements.map((el) => el.type)).toStrictEqual(['text', 'table']);
	});

	it('inserts the staged shape and chart, keeping the picked types', () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		vi.useFakeTimers();
		// A gallery pick stages the type, then inserts it once the host has applied it.
		const pick = (id: string, value: string) => {
			target.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!.click();
			target
				.querySelector<HTMLButtonElement>(
					`[data-ribbon-control="${id}"] [data-insert-item="${value}"]`,
				)!
				.click();
			flushSync();
		};
		pick('insert.illustrations.shapes', 'ellipse');
		pick('insert.illustrations.chart', 'pie');
		vi.runAllTimers();
		vi.useRealTimers();
		flushSync();
		expect(editor.slides[0]?.elements.map((el) => el.type)).toStrictEqual(['shape', 'chart']);
	});

	it('inserts action buttons and fields from the shared menus', () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		control(target, 'insert.text.slideNumber').click();
		target
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.links.action"] .trigger')!
			.click();
		target
			.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.links.action"] [data-insert-item]',
			)!
			.click();
		flushSync();
		expect(editor.slides[0]?.elements.length).toBeGreaterThanOrEqual(1);
	});

	it('arms Freeform tools and reflects the pressed state', () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		const host = target.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!;
		host.click();
		flushSync();
		expect(editor.outlineOps.freeformTool).toBe('curve');
		expect(host.shadowRoot!.querySelector('button')!.getAttribute('aria-pressed')).toBe('true');
	});

	it('inserts a media element from the audio/video file input', async () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		const mediaInput = target.querySelectorAll<HTMLInputElement>('input[type="file"]')[1];
		fireChange(mediaInput, new File(['fake-audio'], 'clip.mp3', { type: 'audio/mpeg' }));
		await vi.waitFor(() => {
			flushSync();
			expect(editor.slides[0]?.elements).toHaveLength(1);
		});
		expect(editor.slides[0]?.elements[0]?.type).toBe('media');
	});

	it('opens the file pickers from the Image and Media commands', () => {
		const target = mountTab(makeEditor());
		const clicks: string[] = [];
		for (const input of target.querySelectorAll<HTMLInputElement>('input[type="file"]')) {
			input.addEventListener('click', (event) => {
				event.preventDefault();
				clicks.push(input.accept);
			});
		}
		pickItem(target, 'insert.images.pictures', 'thisDevice');
		pickItem(target, 'insert.media.media', 'videoThisDevice');
		expect(clicks).toStrictEqual(['image/*', 'video/*,audio/*']);
	});

	it('opens the SmartArt dialog and inserts a layout', async () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		control(target, 'insert.illustrations.smartArt').click();
		flushSync();
		const dialog = document.body.querySelector('[role="dialog"]');
		expect(dialog?.getAttribute('aria-label')).toBe('Insert SmartArt');
		dialog!.querySelector<HTMLButtonElement>('[role="option"]')!.click();
		flushSync();
		allButtons(dialog!)
			.find((button) => button.textContent?.trim() === 'Insert')!
			.click();
		flushSync();
		await tick();
		expect(editor.slides[0]?.elements[0]?.type).toBe('smartArt');
		expect(document.body.querySelector('[role="dialog"]')).toBeNull();
	});

	it('cancels the SmartArt dialog without inserting', async () => {
		const editor = makeEditor();
		const target = mountTab(editor);
		control(target, 'insert.illustrations.smartArt').click();
		flushSync();
		allButtons(document.body.querySelector('[role="dialog"]')!)
			.find((button) => button.textContent?.trim() === 'Cancel')!
			.click();
		flushSync();
		await tick();
		expect(editor.slides[0]?.elements).toHaveLength(0);
		expect(document.body.querySelector('[role="dialog"]')).toBeNull();
	});

	it('toggles the equation dialog open and closed', () => {
		const target = mountTab(makeEditor());
		expect(document.body.querySelector('[role="dialog"]')).toBeNull();
		const face = target.querySelector<HTMLButtonElement>(
			'[data-ribbon-control="insert.symbols.equation"] .primary',
		)!;
		face.click();
		flushSync();
		expect(document.body.querySelector('[role="dialog"]')).not.toBeNull();
		face.click();
		flushSync();
		expect(document.body.querySelector('[role="dialog"]')).toBeNull();
	});

	it('gates controls on editability, Link on the selection, and hides Header & Footer without a handler', () => {
		const editor = makeEditor(false);
		const target = mountTab(editor);
		expect(control(target, 'insert.text.textBox').disabled).toBeTruthy();
		expect(control(target, 'insert.links.link').disabled).toBeTruthy();
		control(target, 'insert.text.textBox').click();
		expect(editor.slides[0]?.elements).toHaveLength(0);
		expect(headerFooter(target).hasAttribute('hidden')).toBeTruthy();
	});

	it('adds a slide from the layout and opens the comments pane', async () => {
		const editor = makeEditor();
		const chromeUi = { commentsOpen: false, toggleComments: vi.fn() };
		const onnavigateslide = vi.fn();
		const target = mountTab(editor, { chromeUi, onnavigateslide });
		await vi.waitFor(() => flushSync());
		target
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.slides.newSlide"] .primary')!
			.click();
		flushSync();
		expect(editor.slides).toHaveLength(2);
		expect(onnavigateslide).toHaveBeenCalledOnce();
		control(target, 'insert.comments.comment').click();
		expect(chromeUi.toggleComments).toHaveBeenCalledOnce();
	});

	it('opens Header & Footer when the host supplies a handler', () => {
		const onheaderfooter = vi.fn();
		const target = mountTab(makeEditor(), { onheaderfooter });
		headerFooter(target).shadowRoot!.querySelector('button')!.click();
		expect(onheaderfooter).toHaveBeenCalledOnce();
	});
});
