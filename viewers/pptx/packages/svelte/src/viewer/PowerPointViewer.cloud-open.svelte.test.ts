/**
 * File > Open > a cloud place, end to end through the real `PowerPointViewer`: the opened file
 * loads like any open, File > Save goes to the file's own storage instead of the host's, and a file
 * the service leaves read-only stays read-only.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type {
	CloudFileProvider,
	CloudOpenedFile,
	DocumentStorage,
	PowerPointViewerAPI,
} from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.svelte';

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');

type ViewerInstance = Pick<
	PowerPointViewerAPI,
	| 'getSlideCount'
	| 'addSlide'
	| 'isDirty'
	| 'addElement'
	| 'updateElement'
	| 'deleteElements'
	| 'getElements'
	| 'getMode'
>;

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	vi.restoreAllMocks();
});

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

async function mountViewer(
	storage: DocumentStorage,
): Promise<{ target: HTMLElement; instance: ViewerInstance }> {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const onload = vi.fn();
	const instance = mount(PowerPointViewer, {
		target,
		props: {
			source: new Uint8Array(readFileSync(FIXTURE)),
			editable: true,
			storage,
			fileProviders: [provider],
			onload,
		},
	}) as unknown as ViewerInstance;
	flushSync();
	cleanup = () => {
		unmount(instance as unknown as Record<string, unknown>);
		target.remove();
	};
	await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
	flushSync();
	return { target, instance };
}

async function cloudFile(
	storage: DocumentStorage,
	canEdit: boolean,
	slides: number,
): Promise<CloudOpenedFile> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slides });
	// One element to try to change, move and delete.
	data.slides[0]!.elements = [
		{ id: 'title-0', type: 'text', text: 'Quarterly report', x: 67, y: 40, width: 400, height: 60 },
	];
	data.slides[0]!.isDirty = true;
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return {
		entry: { id: 'drive!item', name: 'Cloud.pptx', type: 'file', kind: 'pptx' },
		kind: 'pptx',
		fileName: 'Cloud.pptx',
		bytes,
		canEdit,
		storage,
	};
}

function click(element: Element | null | undefined): void {
	if (!element) {
		throw new Error('missing control');
	}
	(element as HTMLElement).click();
	flushSync();
}

/** File tab, Open, the OneDrive place; then the place reports `file` as opened. */
async function openFromCloud(
	target: HTMLElement,
	instance: ViewerInstance,
	file: CloudOpenedFile,
	slides: number,
): Promise<void> {
	click(
		[...target.querySelectorAll('.pptx-svelte-ribbon-tabs [role="tab"]')].find(
			(tab) => tab.textContent?.trim() === 'File',
		),
	);
	click(
		[...target.querySelectorAll('[data-pptx-backstage-nav-item]')].find(
			(item) => item.textContent?.trim() === 'Open',
		),
	);
	target
		.querySelector('pptx-ui-open-places')!
		.dispatchEvent(
			new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
		);
	flushSync();
	// The backstage closed and the cloud deck replaced the host's.
	expect(target.querySelector('pptx-ui-open-places')).toBeNull();
	await vi.waitFor(() => expect(instance.getSlideCount()).toBe(slides), { timeout: 15000 });
	flushSync();
}

/** The title bar's Save command (the shared title bar renders into a shadow root). */
function saveButton(target: HTMLElement): HTMLButtonElement {
	for (const host of target.querySelectorAll('*')) {
		const button = host.shadowRoot?.querySelector<HTMLButtonElement>('[data-command="save"]');
		if (button) {
			return button;
		}
	}
	throw new Error('no Save command rendered');
}

describe('PowerPointViewer opening from a cloud place', () => {
	it('loads the file and saves it to its own storage, not the host storage', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async () => undefined);
		const { target, instance } = await mountViewer({ save: hostSave });
		expect(instance.getSlideCount()).not.toBe(4);
		await openFromCloud(target, instance, await cloudFile({ save: cloudSave }, true, 4), 4);
		instance.addSlide(0);
		flushSync();
		expect(instance.isDirty()).toBe(true);
		saveButton(target).click();
		await vi.waitFor(() => expect(cloudSave).toHaveBeenCalledOnce(), { timeout: 15000 });
		expect(hostSave).not.toHaveBeenCalled();
		await vi.waitFor(() => expect(instance.isDirty()).toBe(false));
	});

	it('stays unsaved when the cloud save fails', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const cloudSave = vi.fn(() => Promise.reject(new Error('offline')));
		const { target, instance } = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(target, instance, await cloudFile({ save: cloudSave }, true, 3), 3);
		instance.addSlide(0);
		flushSync();
		saveButton(target).click();
		await vi.waitFor(() => expect(cloudSave).toHaveBeenCalledOnce(), { timeout: 15000 });
		await Promise.resolve();
		flushSync();
		expect(instance.isDirty()).toBe(true);
	});

	/** The Home tab's New Slide command (a button in the ribbon). */
	function newSlideButton(target: HTMLElement): HTMLButtonElement {
		const button = [...target.querySelectorAll<HTMLButtonElement>('button')].find(
			(candidate) => candidate.getAttribute('aria-label') === 'New Slide',
		);
		if (!button) {
			throw new Error('no New Slide command rendered');
		}
		return button;
	}

	/** Every edit the public handle offers, tried on the active slide. */
	async function tryHandleEdits(instance: ViewerInstance): Promise<void> {
		instance.addElement({ type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never);
		instance.addSlide(0);
		instance.updateElement('title-0', { x: 500 } as never);
		instance.deleteElements(['title-0']);
		flushSync();
		await new Promise((resolve) => {
			setTimeout(resolve, 50);
		});
		flushSync();
	}

	it('opens a file the service leaves read-only: no edit through the handle or the ribbon', async () => {
		const { target, instance } = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(target, instance, await cloudFile({ save: vi.fn() }, false, 3), 3);
		expect(instance.getMode()).toBe('preview');
		const before = instance.getElements(0);
		expect(before).toHaveLength(1);
		expect(
			instance.addElement({ type: 'shape', id: 'p', x: 0, y: 0, width: 5, height: 5 } as never),
		).toBeUndefined();
		await tryHandleEdits(instance);
		expect(instance.getSlideCount()).toBe(3);
		expect(instance.getElements(0)).toStrictEqual(before);
		expect(instance.isDirty()).toBe(false);
		// The ribbon: New Slide is disabled, and clicking it changes nothing.
		expect(newSlideButton(target).disabled).toBe(true);
		click(newSlideButton(target));
		expect(instance.getSlideCount()).toBe(3);
		expect(instance.isDirty()).toBe(false);
	});

	it('edits again when a later cloud file is editable', async () => {
		const { target, instance } = await mountViewer({ save: vi.fn(async () => undefined) });
		await openFromCloud(target, instance, await cloudFile({ save: vi.fn() }, false, 3), 3);
		await tryHandleEdits(instance);
		expect(instance.getSlideCount()).toBe(3);

		await openFromCloud(target, instance, await cloudFile({ save: vi.fn() }, true, 4), 4);
		expect(instance.getMode()).toBe('edit');
		expect(newSlideButton(target).disabled).toBe(false);
		click(newSlideButton(target));
		await vi.waitFor(() => expect(instance.getSlideCount()).toBe(5));
		await tryHandleEdits(instance);
		expect(instance.getSlideCount()).toBe(6);
		expect(instance.isDirty()).toBe(true);
	});
});
