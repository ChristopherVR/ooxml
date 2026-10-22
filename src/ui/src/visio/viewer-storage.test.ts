// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import type {
	DocumentRecoveryEntry,
	DocumentSaveMeta,
	DocumentStorage,
} from 'ooxml-core/embedding';
import type { ViewerController, ViewerState } from './controller';
import { mountViewer } from './binding';
import { ViewerAutoRecover, storageRecoveryStore } from './viewer-autorecover';
import { storageFileName } from './viewer-file-commands';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

function recording(result?: { fileName: string }) {
	const saves: Array<[Uint8Array, DocumentSaveMeta]> = [];
	const storage: DocumentStorage = {
		save: async (bytes, meta) => {
			saves.push([bytes, meta]);
			return result;
		},
	};
	return { storage, saves };
}

async function setup(options: Parameters<typeof mountViewer>[1] = {}) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument, ...options });
	vi.spyOn(viewer.controller, 'parser').mockImplementation(parseVsdx);
	await viewer.createBlankDrawing();
	return { viewer, root: viewer.element.shadowRoot! };
}
const clickSave = (root: ShadowRoot) =>
	root.querySelector<HTMLButtonElement>('[data-backstage-action="download"]')!.click();

describe('<visio-viewer> file-command and host-owned storage', () => {
	it('saves to storage instead of downloading, and marks the drawing saved', async () => {
		const { storage, saves } = recording();
		const { viewer, root } = await setup({ storage });
		expect(viewer.element.storage).toBe(storage);
		const download = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		const commands: string[] = [];
		viewer.element.addEventListener('file-command', (event) =>
			commands.push((event as CustomEvent<{ command: string }>).detail.command),
		);
		viewer.element.openBackstage('save-as');
		clickSave(root);
		await vi.waitFor(() => expect(saves).toHaveLength(1));
		expect(commands).toEqual(['save']);
		expect(download).not.toHaveBeenCalled();
		const [bytes, meta] = saves[0]!;
		expect(bytes.byteLength).toBeGreaterThan(0);
		expect(meta).toEqual({
			fileName: 'New drawing.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			reason: 'save',
		});
		viewer.destroy();
	});

	it('takes the name the host reports back', async () => {
		const { storage } = recording({ fileName: 'Plan.vsdx' });
		const { viewer, root } = await setup({ storage });
		viewer.element.openBackstage('save-as');
		clickSave(root);
		await vi.waitFor(() => expect(viewer.element.fileName).toBe('Plan.vsdx'));
		viewer.destroy();
	});

	it('does nothing when a binding callback returns false or a listener cancels', async () => {
		const { storage, saves } = recording();
		const seen: unknown[] = [];
		const { viewer, root } = await setup({
			storage,
			events: {
				'file-command': (detail) => {
					seen.push(detail);
					return false;
				},
			},
		});
		const download = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		viewer.element.openBackstage('save-as');
		clickSave(root);
		await Promise.resolve();
		expect(seen).toEqual([{ command: 'save' }]);
		expect(saves).toHaveLength(0);
		expect(download).not.toHaveBeenCalled();
		// Without storage a canceled Save does not download either.
		viewer.update({ events: {}, storage: null });
		viewer.element.addEventListener('file-command', (event) => event.preventDefault());
		clickSave(root);
		expect(download).not.toHaveBeenCalled();
		viewer.destroy();
	});

	it('lets a host cancel New, including Ctrl+N', async () => {
		const { viewer, root } = await setup();
		const create = vi.spyOn(viewer.element, 'createBlankDrawing');
		viewer.element.addEventListener('file-command', (event) => event.preventDefault());
		viewer.element.openBackstage('new');
		root.querySelector<HTMLButtonElement>('[data-backstage-action="new-blank"]')!.click();
		root
			.querySelector<HTMLElement>('.viewport')!
			.dispatchEvent(new KeyboardEvent('keydown', { key: 'n', ctrlKey: true, bubbles: true }));
		await Promise.resolve();
		expect(create).not.toHaveBeenCalled();
		viewer.destroy();
	});

	it('opens through storage.open instead of the file picker', async () => {
		const bytes = await createVsdxFixture();
		const click = vi.spyOn(HTMLInputElement.prototype, 'click');
		const { viewer, root } = await setup({
			storage: {
				save: async () => undefined,
				open: async () => ({ bytes, fileName: 'Opened.vsdx' }),
			},
		});
		viewer.element.openBackstage('open');
		root.querySelector<HTMLButtonElement>('[data-backstage-action="open"]')!.click();
		await vi.waitFor(() => expect(viewer.element.fileName).toBe('Opened.vsdx'));
		expect(click).not.toHaveBeenCalled();
		viewer.destroy();
	});

	it("looks for a recovery copy in the storage's recovery store", async () => {
		const readRecovery = vi.fn(async () => undefined);
		const bytes = await createVsdxFixture();
		const { viewer } = await setup({
			storage: {
				save: async () => undefined,
				readRecovery,
				writeRecovery: async () => undefined,
				clearRecovery: async () => undefined,
			},
		});
		await viewer.load(new File([new Uint8Array(bytes)], 'Plan.vsdx'));
		await vi.waitFor(() => expect(readRecovery).toHaveBeenCalledWith('Plan.vsdx'));
		viewer.destroy();
	});
});

describe('<visio-viewer> host saves that fail or race edits', () => {
	const rectangle = (shapeId: string) => ({
		type: 'create-rectangle' as const,
		pageId: '1',
		shapeId,
		x: 2,
		y: 2,
		width: 1,
		height: 1,
	});
	async function editable(storage: DocumentStorage) {
		const { viewer, root } = await setup({ storage });
		type Edit = (bytes: Uint8Array, commands: Parameters<typeof editVsdx>[1]) => Promise<unknown>;
		vi.spyOn(viewer.controller as unknown as { editor: Edit }, 'editor').mockImplementation(
			async (bytes, commands) => {
				const result = await editVsdx(bytes, commands);
				return { ...result, document: await parseVsdx(result.bytes) };
			},
		);
		await viewer.load(new File([new Uint8Array(await createVsdxFixture())], 'Plan.vsdx'));
		await viewer.applyEdits([rectangle('20')]);
		expect(viewer.controller.state.edit.dirty).toBe(true);
		viewer.element.openBackstage('save-as');
		return { viewer, root };
	}
	const deferred = () => {
		let settle!: { resolve: () => void; reject: (error: Error) => void };
		const promise = new Promise<void>((resolve, reject) => (settle = { resolve, reject }));
		return { promise, ...settle };
	};

	it('keeps the drawing dirty and its recovery copy when the host save fails', async () => {
		const clearRecovery = vi.fn(async () => undefined);
		const errors: unknown[] = [];
		const { viewer, root } = await editable({
			save: () => Promise.reject(new Error('offline')),
			readRecovery: async () => undefined,
			writeRecovery: async () => undefined,
			clearRecovery,
		});
		viewer.element.addEventListener('document-error', (event) =>
			errors.push((event as CustomEvent).detail),
		);
		clickSave(root);
		await vi.waitFor(() => expect(errors).toHaveLength(1));
		expect((errors[0] as Error).message).toBe('offline');
		expect(clearRecovery).not.toHaveBeenCalled();
		expect(viewer.controller.state.edit.dirty).toBe(true);
		viewer.destroy();
	});

	it('marks the drawing saved and clears recovery only after the host save succeeds', async () => {
		const pending = deferred();
		const clearRecovery = vi.fn(async () => undefined);
		const save = vi.fn(() => pending.promise);
		const { viewer, root } = await editable({
			save,
			readRecovery: async () => undefined,
			writeRecovery: async () => undefined,
			clearRecovery,
		});
		clickSave(root);
		await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(1));
		expect(clearRecovery).not.toHaveBeenCalled();
		// A second Save while the first is pending is ignored.
		clickSave(root);
		expect(save).toHaveBeenCalledTimes(1);
		pending.resolve();
		await vi.waitFor(() => expect(viewer.controller.state.edit.dirty).toBe(false));
		expect(clearRecovery).toHaveBeenCalledWith('Plan.vsdx');
		viewer.destroy();
	});

	it('stays dirty when the drawing is edited while the host save runs', async () => {
		const pending = deferred();
		const save = vi.fn(() => pending.promise);
		const { viewer, root } = await editable({ save });
		clickSave(root);
		await vi.waitFor(() => expect(save).toHaveBeenCalledTimes(1));
		await viewer.applyEdits([rectangle('21')]);
		pending.resolve();
		await vi.waitFor(() =>
			expect(viewer.element.shadowRoot!.querySelector('[data-status]')!.textContent).toBe(
				'Saved the drawing.',
			),
		);
		expect(viewer.controller.state.edit.dirty).toBe(true);
		viewer.destroy();
	});
});

describe('storageRecoveryStore', () => {
	it('is undefined unless the storage has every recovery method', () => {
		expect(storageRecoveryStore(undefined)).toBeUndefined();
		expect(storageRecoveryStore({ save: async () => undefined })).toBeUndefined();
	});

	it('maps AutoRecover copies onto the recovery methods, keyed by file name', async () => {
		const kept = new Map<string, DocumentRecoveryEntry>();
		const store = storageRecoveryStore({
			save: async () => undefined,
			readRecovery: async (id) => kept.get(id),
			writeRecovery: async (entry) => void kept.set(entry.id, entry),
			clearRecovery: async (id) => void kept.delete(id),
		})!;
		const controller = {
			get state() {
				return {
					document: {},
					loading: false,
					edit: { dirty: true, sourceAvailable: true, busy: false },
				} as unknown as ViewerState;
			},
			exportVsdx: () => ({ bytes: Uint8Array.of(7), dirty: true }),
		} as unknown as ViewerController;
		const timers = { setInterval: () => 1, clearInterval: () => {} };
		const recover = new ViewerAutoRecover(
			controller,
			() => 'Plan.vsdx',
			() => store,
			() => {},
			() => 42,
			timers,
		);
		recover.configure(true, 5);
		expect(await recover.saveNow()).toBe(true);
		expect(kept.get('Plan.vsdx')).toEqual({
			id: 'Plan.vsdx',
			savedAt: 42,
			bytes: Uint8Array.of(7),
		});
		expect(await store.get('Plan.vsdx')).toEqual({
			name: 'Plan.vsdx',
			savedAt: 42,
			bytes: Uint8Array.of(7),
		});
		recover.saved();
		await vi.waitFor(() => expect(kept.size).toBe(0));
	});

	it('names saved drawings as .vsdx', () => {
		expect(storageFileName('Plan.vsd')).toBe('Plan.vsdx');
		expect(storageFileName('')).toBe('Drawing1.vsdx');
	});
});
