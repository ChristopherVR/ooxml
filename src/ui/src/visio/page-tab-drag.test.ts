import { afterEach, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { mountViewer } from './binding';
import type { EditWorkerRequest } from './worker-editor';
import { fixture } from '../../../core/visio/test-fixtures';

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
afterEach(() => {
	vi.unstubAllGlobals();
	document.body.replaceChildren();
});

const pointer = (target: Element, type: string, clientX: number) => {
	const event = new MouseEvent(type, { bubbles: true, clientX, button: 0 });
	Object.defineProperty(event, 'pointerId', { value: 1 });
	target.dispatchEvent(event);
};

async function open(background = false) {
	vi.stubGlobal('Worker', CoreWorker);
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host);
	const pages: { id: string; contents: string; attributes?: string }[] = ['0', '1', '2'].map(
		(id) => ({ id, contents: '<Shapes/>' }),
	);
	if (background) pages.push({ id: '3', contents: '<Shapes/>', attributes: 'Background="1"' });
	await viewer.load(await fixture({ pages }));
	const strip = viewer.element.shadowRoot!.querySelector('office-ui-tab-strip')!;
	const tabs = () => [...strip.shadowRoot!.querySelectorAll<HTMLElement>('[role="tab"]')];
	tabs().forEach((t, i) => {
		t.getBoundingClientRect = () =>
			({
				left: i * 100,
				right: i * 100 + 100,
				width: 100,
				top: 0,
				bottom: 30,
				height: 30,
			}) as DOMRect;
	});
	const order = () => viewer.element.document!.pages.map((page) => page.id);
	return { viewer, strip, tabs, order };
}

it('drags a page tab to a new place as one undoable edit that survives save', async () => {
	const { viewer, strip, tabs, order } = await open();
	try {
		expect(strip.hasAttribute('reorderable')).toBe(true);
		pointer(tabs()[0]!, 'pointerdown', 10);
		pointer(tabs()[0]!, 'pointermove', 240);
		expect(strip.shadowRoot!.querySelector('.marker')).not.toBeNull();
		pointer(tabs()[0]!, 'pointerup', 240);
		expect(strip.shadowRoot!.querySelector('.marker')).toBeNull();
		await vi.waitFor(() => expect(order()).toEqual(['1', '0', '2']));
		expect((await parseVsdx(viewer.exportVsdx().bytes)).pages.map((page) => page.id)).toEqual([
			'1',
			'0',
			'2',
		]);
		await viewer.undo();
		expect(order()).toEqual(['0', '1', '2']);
		await viewer.redo();
		expect(order()).toEqual(['1', '0', '2']);
	} finally {
		viewer.destroy();
	}
});

it('keeps a click as page selection and cancels a drag with Escape', async () => {
	const { viewer, tabs, order } = await open();
	try {
		pointer(tabs()[1]!, 'pointerdown', 150);
		pointer(tabs()[1]!, 'pointermove', 152);
		pointer(tabs()[1]!, 'pointerup', 152);
		tabs()[1]!.click();
		expect(viewer.element.pageIndex).toBe(1);
		pointer(tabs()[0]!, 'pointerdown', 10);
		pointer(tabs()[0]!, 'pointermove', 250);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		pointer(tabs()[0]!, 'pointerup', 250);
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(order()).toEqual(['0', '1', '2']);
	} finally {
		viewer.destroy();
	}
});

it('does not drag a page past the background pages', async () => {
	const { viewer, tabs, order } = await open(true);
	try {
		pointer(tabs()[0]!, 'pointerdown', 10);
		pointer(tabs()[0]!, 'pointermove', 390);
		pointer(tabs()[0]!, 'pointerup', 390);
		await vi.waitFor(() => expect(order()).toEqual(['1', '2', '0', '3']));
		pointer(tabs()[3]!, 'pointerdown', 310);
		pointer(tabs()[3]!, 'pointermove', 10);
		pointer(tabs()[3]!, 'pointerup', 10);
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(order()).toEqual(['1', '2', '0', '3']);
	} finally {
		viewer.destroy();
	}
});
