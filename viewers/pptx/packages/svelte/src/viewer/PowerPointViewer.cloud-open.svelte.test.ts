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
	'getSlideCount' | 'addSlide' | 'isDirty' | 'addElement'
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
	click(target.querySelector('[data-pptx-open-place="cloud:onedrive"]'));
	target
		.querySelector('[data-pptx-cloud-place] pptx-ui-cloud-files')!
		.dispatchEvent(
			new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
		);
	flushSync();
	// The backstage closed and the cloud deck replaced the host's.
	expect(target.querySelector('[data-pptx-cloud-place]')).toBeNull();
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

	it('opens a file the service leaves read-only without letting it be edited', async () => {
		const { target, instance } = await mountViewer({ save: vi.fn(async () => undefined) });
		const shape = { type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never;
		expect(instance.addElement(shape)).toBeDefined();
		await openFromCloud(target, instance, await cloudFile({ save: vi.fn() }, false, 3), 3);
		expect(instance.addElement(shape)).toBeUndefined();
	});
});
