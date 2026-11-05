/**
 * File > Open > a cloud place, end to end through the real `PowerPointViewer.vue`: the opened file
 * loads like any open, File > Save goes to the file's own storage instead of the host's, and a file
 * the service leaves read-only stays read-only.
 */
import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import type { CloudFileProvider, CloudOpenedFile, DocumentStorage } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

async function sampleDeck(slideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	// One element to try to change, move and delete.
	data.slides[0]!.elements = [
		{ id: 'title-0', type: 'text', text: 'Quarterly report', x: 67, y: 40, width: 400, height: 60 },
	];
	data.slides[0]!.isDirty = true;
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
	vi.restoreAllMocks();
});

async function settle(done: () => boolean = () => false, attempts = 20): Promise<void> {
	for (let attempt = 0; attempt < attempts && !done(); attempt += 1) {
		await flushPromises();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
}

async function mountViewer(storage: DocumentStorage): Promise<PowerPointViewerExpose> {
	const wrapper = mount(PowerPointViewer, {
		props: {
			content: await sampleDeck(2),
			canEdit: true,
			storage,
			fileProviders: [provider],
		},
		attachTo: document.body,
	});
	mounted = wrapper;
	const viewer = wrapper.vm as unknown as PowerPointViewerExpose;
	await settle(() => viewer.getSlideCount() > 0, 100);
	if (viewer.getSlideCount() === 0) {
		throw new Error('the viewer never finished loading its deck');
	}
	return viewer;
}

async function click(element: Element | null | undefined): Promise<void> {
	if (!element) {
		throw new Error('missing control');
	}
	(element as HTMLElement).click();
	await flushPromises();
}

/** File tab, Open, the OneDrive place; then the place reports `file` as opened. */
async function openFromCloud(file: CloudOpenedFile): Promise<void> {
	const root = mounted!.element.parentElement ?? document.body;
	await click(
		[...root.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent?.trim() === 'File'),
	);
	await click(
		[...document.querySelectorAll('[data-pptx-backstage-nav-item]')].find(
			(item) => item.textContent?.trim() === 'Open',
		),
	);
	document
		.querySelector('pptx-ui-open-places')!
		.dispatchEvent(
			new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
		);
	await settle();
}

/** The title bar's Save command (the quick-access toolbar lives in a shadow root). */
function saveButton(): HTMLButtonElement {
	for (const host of document.querySelectorAll('*')) {
		const button = host.shadowRoot?.querySelector<HTMLButtonElement>('[data-command="save"]');
		if (button) {
			return button;
		}
	}
	throw new Error('no Save command rendered');
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
		bytes: await sampleDeck(slides),
		canEdit,
		storage,
	};
}

describe('powerPointViewer opening from a cloud place', () => {
	it('loads the file and saves it to its own storage, not the host storage', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async () => undefined);
		const viewer = await mountViewer({ save: hostSave });
		expect(viewer.getSlideCount()).toBe(2);
		await openFromCloud(await cloudFile({ save: cloudSave }, true, 4));
		await settle(() => viewer.getSlideCount() === 4, 100);
		expect(viewer.getSlideCount()).toBe(4);
		expect(document.querySelector('pptx-ui-open-places')).toBeNull();
		viewer.addSlide(0);
		await flushPromises();
		expect(viewer.isDirty()).toBe(true);
		saveButton().click();
		await settle(() => cloudSave.mock.calls.length > 0 && !viewer.isDirty());
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(hostSave).not.toHaveBeenCalled();
		expect(viewer.isDirty()).toBe(false);
	});

	it('stays unsaved when the cloud save fails', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const cloudSave = vi.fn(() => Promise.reject(new Error('offline')));
		const viewer = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: cloudSave }, true, 3));
		await settle(() => viewer.getSlideCount() === 3, 100);
		viewer.addSlide(0);
		await flushPromises();
		saveButton().click();
		await settle(() => cloudSave.mock.calls.length > 0);
		await settle();
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(viewer.isDirty()).toBe(true);
	});

	/** The Home tab's New Slide command (a button in the ribbon). */
	function newSlideButton(): HTMLButtonElement {
		const button = [...document.querySelectorAll<HTMLButtonElement>('button')].find(
			(candidate) => candidate.getAttribute('aria-label') === 'New Slide',
		);
		if (!button) {
			throw new Error('no New Slide command rendered');
		}
		return button;
	}

	/** Every edit the public handle offers, tried on the active slide. */
	async function tryHandleEdits(viewer: PowerPointViewerExpose): Promise<void> {
		viewer.addElement({ type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never);
		viewer.addSlide(0);
		viewer.updateElement('title-0', { x: 500 } as never);
		viewer.deleteElements(['title-0']);
		await settle();
	}

	it('opens a file the service leaves read-only: no edit through the handle or the ribbon', async () => {
		const viewer = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: vi.fn() }, false, 3));
		await settle(() => viewer.getSlideCount() === 3, 100);
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.getMode()).toBe('preview');
		const before = viewer.getElements(0);
		expect(before).toHaveLength(1);
		expect(
			viewer.addElement({ type: 'shape', id: 'p', x: 0, y: 0, width: 5, height: 5 } as never),
		).toBeUndefined();
		await tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.getElements(0)).toStrictEqual(before);
		expect(viewer.isDirty()).toBe(false);
		// The ribbon: New Slide is disabled, and clicking it changes nothing.
		expect(newSlideButton().disabled).toBe(true);
		await click(newSlideButton());
		await settle();
		expect(viewer.getSlideCount()).toBe(3);
		expect(viewer.isDirty()).toBe(false);
	});

	it('edits again when a later cloud file is editable', async () => {
		const viewer = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: vi.fn() }, false, 3));
		await settle(() => viewer.getSlideCount() === 3, 100);
		await tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(3);

		await openFromCloud(await cloudFile({ save: vi.fn() }, true, 4));
		await settle(() => viewer.getSlideCount() === 4, 100);
		expect(viewer.getSlideCount()).toBe(4);
		expect(viewer.getMode()).toBe('edit');
		expect(newSlideButton().disabled).toBe(false);
		await click(newSlideButton());
		await settle();
		expect(viewer.getSlideCount()).toBe(5);
		await tryHandleEdits(viewer);
		expect(viewer.getSlideCount()).toBe(6);
		expect(viewer.isDirty()).toBe(true);
	});
});
