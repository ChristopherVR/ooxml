import { afterEach, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { mountViewer } from './binding';
import type { EditWorkerRequest } from './worker-editor';
import { cell, fixture } from '../../../core/visio/test-fixtures';

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

it('inserts from the page bar, saves, selects the new page and restores history without an invalid index', async () => {
	vi.stubGlobal('Worker', CoreWorker);
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host);
	try {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: '<Shapes/>',
					pageCells: cell('DrawingScale', 2) + cell('PageScale', 1),
				},
			],
		});
		await viewer.load(bytes);
		const tabs = viewer.element.shadowRoot!.querySelector('office-ui-tab-strip')!;
		const add = () => tabs.shadowRoot!.querySelector<HTMLButtonElement>('.add')!;
		expect(add().disabled).toBe(false);
		add().click();
		await vi.waitFor(() => expect(viewer.element.document!.pages).toHaveLength(2));
		expect(viewer.element.pageIndex).toBe(1);
		expect(viewer.element.document!.pages[1]).toMatchObject({
			id: '1',
			name: 'Page-2',
			width: 4.25,
			height: 5.5,
			shapes: [],
		});
		const saved = await parseVsdx(viewer.exportVsdx().bytes);
		expect(saved.pages).toHaveLength(2);
		await viewer.undo();
		expect(viewer.element.document!.pages).toHaveLength(1);
		expect(viewer.element.pageIndex).toBe(0);
		expect(add().disabled).toBe(false);
		await viewer.redo();
		expect(viewer.element.document!.pages).toHaveLength(2);
		expect(tabs.shadowRoot!.querySelectorAll('[role="tab"]')).toHaveLength(2);
	} finally {
		viewer.destroy();
	}
});
