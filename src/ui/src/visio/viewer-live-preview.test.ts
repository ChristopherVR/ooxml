import { afterEach, describe, expect, it, vi } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { mountViewer } from './binding';
import type { EditWorkerRequest } from './worker-editor';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';
import { ViewerLivePreview } from './viewer-live-preview';

afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

const pick = (gallery: Element, itemId: string) =>
	gallery.dispatchEvent(
		new CustomEvent('office-gallery-pick', { detail: { gallery: 'quick-styles', itemId } }),
	);
/** The id of the n-th tile of the Quick Styles gallery, so a test never guesses at valid ids. */
const tile = (gallery: Element, n = 0) =>
	(gallery as unknown as { state: { sections: { items: { id: string }[] }[] } }).state.sections[0]!
		.items[n]!.id;
const shapeStyle = (document: VisioDocument | null) =>
	JSON.stringify(document!.pages[0]!.shapes[0]!.style);

describe('controller previews', () => {
	it('shows the drawing an edit would produce without committing it', async () => {
		const view = await setupFormattingViewer();
		view.selection();
		const before = view.controller.state.document!;
		const gallery = view.root.querySelector('[data-menu="quick-styles"]')!;
		expect(gallery.hasAttribute('disabled')).toBe(false);
		const shown: (VisioDocument | undefined)[] = [];
		view.controller.previewing(
			(document) => shown.push(document),
			() => pick(gallery, tile(gallery)),
		);
		await vi.waitFor(() => expect(shown).toHaveLength(1));
		expect(shown[0]).toBeDefined();
		expect(shapeStyle(shown[0]!)).not.toBe(shapeStyle(before));
		// Nothing was committed: same drawing object, no history, not dirty, no announcement.
		const state = view.controller.state;
		expect(state.document).toBe(before);
		expect(state.edit.canUndo).toBe(false);
		expect(state.edit.dirty).toBe(false);
		expect(view.feedback).toEqual([]);
		view.dispose();
	});

	it('asks the editor for exactly the edit the real pick would apply', async () => {
		const view = await setupFormattingViewer();
		view.selection();
		const gallery = view.root.querySelector('[data-menu="quick-styles"]')!;
		const id = tile(gallery, 9);
		const [, color, matrix] = /^quick-style-(\d+)-(\d+)$/.exec(id)!;
		const shown: unknown[] = [];
		view.controller.previewing(
			(document) => shown.push(document),
			() => pick(gallery, id),
		);
		await vi.waitFor(() => expect(shown).toHaveLength(1));
		const previewed = view.edits.at(-1)!;
		expect(previewed).toHaveLength(1);
		expect(previewed[0]).toMatchObject({
			type: 'format-shape',
			quickStyle: { color: Number(color), matrix: Number(matrix) },
		});
		view.dispose();
	});

	it('reports being in a preview only while one is being asked for', async () => {
		const view = await setupFormattingViewer();
		expect(view.controller.isPreviewing).toBe(false);
		let during = false;
		view.controller.previewing(
			() => {},
			() => {
				during = view.controller.isPreviewing;
			},
		);
		expect(during).toBe(true);
		expect(view.controller.isPreviewing).toBe(false);
		view.dispose();
	});

	it('discards a preview that a real edit overtook', async () => {
		const view = await setupFormattingViewer();
		view.selection();
		const gallery = view.root.querySelector('[data-menu="quick-styles"]')!;
		const shown: (VisioDocument | undefined)[] = [];
		view.controller.previewing(
			(document) => shown.push(document),
			() => pick(gallery, tile(gallery)),
		);
		view.press('bold');
		await view.done();
		await vi.waitFor(() => expect(shown).toHaveLength(1));
		expect(shown[0]).toBeUndefined();
		expect(view.controller.state.edit.canUndo).toBe(true);
		view.dispose();
	});

	it('discards a preview that was cancelled', async () => {
		const view = await setupFormattingViewer();
		view.selection();
		const gallery = view.root.querySelector('[data-menu="quick-styles"]')!;
		const shown: (VisioDocument | undefined)[] = [];
		view.controller.previewing(
			(document) => shown.push(document),
			() => pick(gallery, tile(gallery)),
		);
		view.controller.cancelPreview();
		await vi.waitFor(() => expect(shown).toHaveLength(1));
		expect(shown[0]).toBeUndefined();
		view.dispose();
	});

	it('reports an edit the editor would refuse as no preview', async () => {
		const view = await setupFormattingViewer();
		view.selection();
		const gallery = view.root.querySelector('[data-menu="quick-styles"]')!;
		const shown: (VisioDocument | undefined)[] = [];
		view.controller.previewing(
			(document) => shown.push(document),
			() => pick(gallery, 'quick-style-1-99'),
		);
		await vi.waitFor(() => expect(shown).toEqual([undefined]));
		expect(view.controller.state.edit.error).toBeNull();
		view.dispose();
	});

	it('does nothing for a read-only drawing', async () => {
		const view = await setupFormattingViewer(false);
		const shown: (VisioDocument | undefined)[] = [];
		await view.controller.previewing(
			(document) => shown.push(document),
			() => void view.controller.applyEdits([]).catch(() => {}),
		);
		await vi.waitFor(() => expect(shown).toEqual([undefined]));
		view.dispose();
	});
});

describe('hovering a gallery', () => {
	function rig() {
		const host = document.createElement('div');
		document.body.append(host);
		const root = host.attachShadow({ mode: 'open' });
		const gallery = document.createElement('office-ui-gallery');
		root.append(gallery);
		const picked: string[] = [];
		gallery.addEventListener('office-gallery-pick', (event) =>
			picked.push((event as CustomEvent<{ itemId: string }>).detail.itemId),
		);
		const controller = {
			previewing: vi.fn((_sink: (document: VisioDocument | undefined) => void, apply: () => void) =>
				apply(),
			),
			cancelPreview: vi.fn(),
		};
		const show = vi.fn();
		const end = vi.fn();
		const preview = new ViewerLivePreview(root, controller as never, show, end);
		const dispose = preview.wire();
		const hover = (itemId: string) =>
			gallery.dispatchEvent(
				new CustomEvent('office-gallery-preview', {
					detail: { gallery: 'quick-styles', itemId },
					bubbles: true,
					composed: true,
				}),
			);
		const leave = () =>
			gallery.dispatchEvent(
				new CustomEvent('office-gallery-preview-end', { bubbles: true, composed: true }),
			);
		return { preview, controller, show, end, picked, hover, leave, dispose };
	}

	it('previews the tile the pointer rests on, not each one it passes', () => {
		vi.useFakeTimers();
		const { controller, picked, hover } = rig();
		hover('one');
		vi.advanceTimersByTime(50);
		hover('two');
		vi.advanceTimersByTime(50);
		hover('three');
		expect(controller.previewing).not.toHaveBeenCalled();
		vi.advanceTimersByTime(200);
		expect(controller.previewing).toHaveBeenCalledTimes(1);
		expect(picked).toEqual(['three']);
	});

	it('shows the drawing it is given and goes back when the tile is left', () => {
		vi.useFakeTimers();
		const { controller, show, end, hover, leave } = rig();
		const document = { pages: [] } as unknown as VisioDocument;
		controller.previewing.mockImplementation(
			(sink: (document: VisioDocument | undefined) => void, apply: () => void) => {
				apply();
				sink(document);
			},
		);
		hover('one');
		vi.advanceTimersByTime(200);
		expect(show).toHaveBeenCalledWith(document);
		leave();
		expect(controller.cancelPreview).toHaveBeenCalled();
		expect(end).toHaveBeenCalled();
	});

	it('shows nothing for a refused preview and nothing once the tile is left', () => {
		vi.useFakeTimers();
		const { controller, show, hover, leave } = rig();
		controller.previewing.mockImplementation(
			(sink: (document: VisioDocument | undefined) => void) => sink(undefined),
		);
		hover('one');
		vi.advanceTimersByTime(200);
		expect(show).not.toHaveBeenCalled();
		leave();
		vi.advanceTimersByTime(500);
		expect(controller.previewing).toHaveBeenCalledTimes(1);
	});

	it('does nothing when Live Preview is off, and cancels a waiting preview when it is', () => {
		vi.useFakeTimers();
		const { preview, controller, hover } = rig();
		preview.enabled = false;
		hover('one');
		vi.advanceTimersByTime(300);
		expect(controller.previewing).not.toHaveBeenCalled();
		preview.enabled = true;
		hover('two');
		preview.stop();
		vi.advanceTimersByTime(300);
		expect(controller.previewing).not.toHaveBeenCalled();
	});

	it('stops listening once disposed', () => {
		vi.useFakeTimers();
		const { controller, hover, dispose } = rig();
		dispose();
		hover('one');
		vi.advanceTimersByTime(300);
		expect(controller.previewing).not.toHaveBeenCalled();
	});
});

/** Transport shim exercises the actual core backend without browser worker scheduling. */
class CoreWorker {
	onmessage: ((event: MessageEvent) => void) | null = null;
	onerror: ((event: ErrorEvent) => void) | null = null;
	#terminated = false;
	terminate(): void {
		this.#terminated = true;
	}
	postMessage(request: ArrayBuffer | EditWorkerRequest): void {
		void (async () => {
			try {
				const result =
					request instanceof ArrayBuffer
						? { document: await parseVsdx(request) }
						: await (async () => {
								const saved = await editVsdx(request.bytes, request.edits);
								return { ...saved, document: await parseVsdx(saved.bytes) };
							})();
				if (!this.#terminated) this.onmessage?.({ data: { ok: true, ...result } } as MessageEvent);
			} catch (error) {
				if (!this.#terminated)
					this.onmessage?.({ data: { ok: false, message: String(error) } } as MessageEvent);
			}
		})();
	}
}

describe('Live Preview in the viewer', () => {
	async function viewer() {
		vi.stubGlobal('Worker', CoreWorker);
		localStorage.clear();
		const host = document.createElement('div');
		document.body.append(host);
		const mounted = mountViewer(host);
		await mounted.load(await createVsdxFixture('Formatted shape'));
		mounted.element.controller.selectShape({ id: '1', name: 'Import test', pageId: '1' });
		const root = mounted.element.shadowRoot!;
		const gallery = root.querySelector('[data-menu="quick-styles"]')!;
		const paint = () => root.querySelector('[data-shape-id="1"] path')?.outerHTML ?? '';
		const hover = (n = 9) =>
			gallery.dispatchEvent(
				new CustomEvent('office-gallery-preview', {
					detail: { gallery: 'quick-styles', itemId: tile(gallery, n) },
					bubbles: true,
					composed: true,
				}),
			);
		const leave = () =>
			gallery.dispatchEvent(
				new CustomEvent('office-gallery-preview-end', { bubbles: true, composed: true }),
			);
		return { mounted, paint, hover, leave };
	}

	it('draws the hovered style on the shape, then puts the shape back', async () => {
		const { mounted, paint, hover, leave } = await viewer();
		const before = paint();
		expect(before).not.toBe('');
		hover();
		await vi.waitFor(() => expect(paint()).not.toBe(before), { timeout: 4000 });
		// Only the picture changed: nothing to undo, nothing saved.
		const state = mounted.element.controller.state;
		expect(state.edit.canUndo).toBe(false);
		expect(state.edit.dirty).toBe(false);
		leave();
		expect(paint()).toBe(before);
		mounted.destroy();
	});

	it('shows nothing when the Live Preview option is off', async () => {
		vi.stubGlobal('Worker', CoreWorker);
		localStorage.setItem('ooxml-visio-options', JSON.stringify({ livePreview: false }));
		const host = document.createElement('div');
		document.body.append(host);
		const mounted = mountViewer(host);
		await mounted.load(await createVsdxFixture('Formatted shape'));
		mounted.element.controller.selectShape({ id: '1', name: 'Import test', pageId: '1' });
		const root = mounted.element.shadowRoot!;
		const gallery = root.querySelector('[data-menu="quick-styles"]')!;
		const before = root.querySelector('[data-shape-id="1"] path')?.outerHTML;
		gallery.dispatchEvent(
			new CustomEvent('office-gallery-preview', {
				detail: { gallery: 'quick-styles', itemId: tile(gallery, 9) },
				bubbles: true,
				composed: true,
			}),
		);
		await new Promise((resolve) => setTimeout(resolve, 600));
		expect(root.querySelector('[data-shape-id="1"] path')?.outerHTML).toBe(before);
		mounted.destroy();
	});
});
