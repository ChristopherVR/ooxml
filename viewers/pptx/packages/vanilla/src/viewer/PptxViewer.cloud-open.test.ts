/**
 * File > Open > a cloud place, end to end through the real vanilla viewer: the opened file loads
 * like any open, File > Save goes to the file's own storage instead of the host's, and a file the
 * service leaves read-only stays read-only until another document is loaded.
 */
import type { CloudFileProvider, CloudOpenedFile, DocumentStorage } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer, PptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

let active: PptxViewerInstance[] = [];

async function deck(slides: number): Promise<Uint8Array> {
	const { PptxHandler } = await import('pptx-viewer-core');
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slides });
	// One element to try to change, move and delete.
	data.slides[0]!.elements = [
		{ id: 'title-0', type: 'text', text: 'Quarterly report', x: 67, y: 40, width: 400, height: 60 },
	];
	data.slides[0]!.isDirty = true;
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

async function mount(storage: DocumentStorage): Promise<{ viewer: PptxViewer; root: HTMLElement }> {
	const root = document.createElement('div');
	document.body.appendChild(root);
	const viewer = createPptxViewer(root, {
		storage,
		fileName: 'Host.pptx',
		editable: true,
		fileProviders: [provider],
	});
	active.push(viewer);
	await viewer.loadFile(await deck(2));
	return { viewer: viewer as PptxViewer, root };
}

async function cloudFile(
	storage: DocumentStorage,
	canEdit: boolean,
	slides: number,
): Promise<CloudOpenedFile> {
	return {
		entry: { id: 'drive!item', name: 'Cloud.pptx', type: 'file', kind: 'pptx' },
		kind: 'pptx',
		fileName: 'Cloud.pptx',
		bytes: await deck(slides),
		canEdit,
		storage,
	};
}

async function until(check: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !check(); attempt += 1) {
		await new Promise((resolve) => {
			setTimeout(resolve, 10);
		});
	}
}

/** File, Open, the OneDrive place; then the place reports `file` as opened. */
async function openFromCloud(
	viewer: PptxViewer,
	root: HTMLElement,
	file: CloudOpenedFile,
	slides: number,
): Promise<void> {
	const backstage = root.querySelector<HTMLElement>('.pptxv-backstage')!;
	[...backstage.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]')]
		.find((item) => item.textContent === translationsEn['pptx.backstage.nav.open'])
		?.click();
	const browser = backstage.querySelector('pptx-ui-open-places');
	expect(browser).not.toBeNull();
	browser?.dispatchEvent(
		new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
	);
	await until(() => viewer.getSlideCount() === slides);
}

afterEach(async () => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
	vi.restoreAllMocks();
	await new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});
});

describe('PptxViewer opening from a cloud place', () => {
	it('loads the file and saves it to its own storage, not the host storage', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async (_bytes: Uint8Array, _meta: unknown) => undefined);
		const { viewer, root } = await mount({ save: hostSave });
		expect(viewer.getSlideCount()).toBe(2);
		const file = await cloudFile({ save: cloudSave }, true, 4);
		await openFromCloud(viewer, root, file, 4);
		expect(viewer.getSlideCount()).toBe(4);
		expect(viewer.getFileName()).toBe('Cloud.pptx');
		viewer.addSlide(0);
		expect(viewer.isDirty()).toBe(true);
		await viewer.downloadPptx();
		expect(cloudSave).toHaveBeenCalledWith(
			expect.any(Uint8Array),
			expect.objectContaining({ reason: 'save' }),
		);
		expect(hostSave).not.toHaveBeenCalled();
		expect(viewer.isDirty()).toBe(false);
	});

	it('stays unsaved when the cloud save fails', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const cloudSave = vi.fn(() => Promise.reject(new Error('offline')));
		const { viewer, root } = await mount({ save: vi.fn(async () => undefined) });
		const file = await cloudFile({ save: cloudSave }, true, 3);
		await openFromCloud(viewer, root, file, 3);
		viewer.addSlide(0);
		await viewer.downloadPptx();
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(viewer.isDirty()).toBe(true);
		expect(viewer.store.get().compatToasts.map((toast) => toast.id)).toContain('save-failed');
	});

	it('keeps a file the service leaves read-only from being edited, until another deck loads', async () => {
		const hostSave = vi.fn(async () => undefined);
		const { viewer, root } = await mount({ save: hostSave });
		const shape = { type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never;
		expect(viewer.addElement(shape)).toBeDefined();
		const file = await cloudFile({ save: vi.fn() }, false, 3);
		await openFromCloud(viewer, root, file, 3);
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.addElement(shape)).toBeUndefined();
		// The host loads its own deck: its storage, name and edit right return.
		await viewer.loadFile(await deck(2));
		expect(viewer.getFileName()).toBe('Host.pptx');
		expect(viewer.addElement(shape)).toBeDefined();
		await viewer.downloadPptx();
		expect(hostSave).toHaveBeenCalledOnce();
	});

	/** The Home tab's New Slide command (a ribbon command element wrapping a button). */
	function newSlideButton(root: HTMLElement): HTMLButtonElement {
		const direct = [...root.querySelectorAll<HTMLButtonElement>('button')].find(
			(candidate) => candidate.getAttribute('aria-label') === 'New Slide',
		);
		const command = [...root.querySelectorAll('pptx-ui-ribbon-command')]
			.filter((host) => host.getAttribute('label') === 'New Slide')
			.map((host) => host.shadowRoot?.querySelector<HTMLButtonElement>('button'))[0];
		const button = direct ?? command;
		if (!button) {
			throw new Error('no New Slide command rendered');
		}
		return button;
	}

	/** Every edit the public handle offers, tried on the active slide. */
	function tryHandleEdits(viewer: PptxViewer): void {
		viewer.addElement({ type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never);
		viewer.addSlide(0);
		viewer.updateElement('title-0', { x: 500 } as never);
		viewer.deleteElements(['title-0']);
	}

	it('opens a file the service leaves read-only: no edit through the handle or the ribbon', async () => {
		const { viewer, root } = await mount({ save: vi.fn(async () => undefined) });
		await openFromCloud(viewer, root, await cloudFile({ save: vi.fn() }, false, 3), 3);
		expect(viewer.getMode()).toBe('preview');
		const before = viewer.getElements(0);
		expect(before).toHaveLength(1);
		expect(
			viewer.addElement({ type: 'shape', id: 'p', x: 0, y: 0, width: 5, height: 5 } as never),
		).toBeUndefined();
		tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.getElements(0)).toStrictEqual(before);
		expect(viewer.isDirty()).toBe(false);
		// The ribbon: New Slide is disabled, and clicking it changes nothing.
		expect(newSlideButton(root).disabled).toBe(true);
		newSlideButton(root).click();
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.isDirty()).toBe(false);
	});

	it('edits again when a later cloud file is editable', async () => {
		const { viewer, root } = await mount({ save: vi.fn(async () => undefined) });
		await openFromCloud(viewer, root, await cloudFile({ save: vi.fn() }, false, 3), 3);
		tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(3);

		await openFromCloud(viewer, root, await cloudFile({ save: vi.fn() }, true, 4), 4);
		expect(viewer.getMode()).toBe('edit');
		expect(newSlideButton(root).disabled).toBe(false);
		newSlideButton(root).click();
		expect(viewer.getSlideCount()).toBe(5);
		tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(6);
		expect(viewer.isDirty()).toBe(true);
	});
});
