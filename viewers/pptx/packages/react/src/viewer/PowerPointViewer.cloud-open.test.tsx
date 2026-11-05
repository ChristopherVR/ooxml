// @vitest-environment happy-dom
/**
 * File > Open > a cloud place, end to end through the real `PowerPointViewer`: the opened file
 * loads like any open, File > Save goes to the file's own storage instead of the host's, and a file
 * the service leaves read-only stays read-only.
 */
import type { CloudFileProvider, CloudOpenedFile, DocumentStorage } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../index');
type ViewerHandle = import('../index').PowerPointViewerHandle;

async function sampleDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	vi.restoreAllMocks();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

async function settle(): Promise<void> {
	await act(async () => {
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	});
}

/**
 * The live handle: `useImperativeHandle` publishes a new handle per render, so it is re-read on
 * every call instead of kept.
 */
type LiveHandle = () => ViewerHandle;

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

async function mount(storage: DocumentStorage): Promise<LiveHandle> {
	const bytes = await sampleDeck();
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(
			<PowerPointViewer
				ref={ref}
				content={bytes}
				canEdit
				storage={storage}
				fileProviders={[provider]}
			/>,
		);
	});
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0) {
			return () => ref.current!;
		}
		await settle();
	}
	throw new Error('the viewer never finished loading its deck');
}

/** The title bar's Save command (the quick-access toolbar lives in a shadow root). */
function saveButton(): HTMLButtonElement {
	for (const host of container.querySelectorAll('*')) {
		const button = host.shadowRoot?.querySelector<HTMLButtonElement>('[data-command="save"]');
		if (button) {
			return button;
		}
	}
	throw new Error('no Save command rendered');
}

function click(element: Element | null | undefined): void {
	if (!element) {
		throw new Error('missing control');
	}
	act(() => (element as HTMLElement).click());
}

/** File tab, Open; then the place reports `file` as opened. */
async function openFromCloud(file: CloudOpenedFile): Promise<void> {
	click([...container.querySelectorAll('[role="tab"]')].find((tab) => tab.textContent === 'File'));
	click(
		[...container.querySelectorAll('[data-pptx-backstage-nav-item]')].find(
			(item) => item.textContent === translationsEn['pptx.backstage.nav.open'],
		),
	);
	await act(async () => {
		container
			.querySelector('pptx-ui-open-places')!
			.dispatchEvent(
				new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
			);
	});
	for (let attempt = 0; attempt < 20; attempt += 1) {
		await settle();
	}
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

describe('PowerPointViewer opening from a cloud place', () => {
	it('loads the file and saves it to its own storage, not the host storage', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async () => undefined);
		const viewer = await mount({ save: hostSave });
		expect(viewer().getSlideCount()).toBe(2);
		await openFromCloud(await cloudFile({ save: cloudSave }, true, 4));
		expect(viewer().getSlideCount()).toBe(4);
		await act(async () => {
			viewer().addSlide(0);
			await Promise.resolve();
		});
		expect(viewer().isDirty()).toBe(true);
		await act(async () => {
			saveButton().click();
		});
		for (let attempt = 0; attempt < 20; attempt += 1) {
			await settle();
		}
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(hostSave).not.toHaveBeenCalled();
		expect(viewer().isDirty()).toBe(false);
	});

	it('stays unsaved when the cloud save fails', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const cloudSave = vi.fn(() => Promise.reject(new Error('offline')));
		const viewer = await mount({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: cloudSave }, true, 3));
		await act(async () => {
			viewer().addSlide(0);
			await Promise.resolve();
		});
		await act(async () => {
			saveButton().click();
		});
		for (let attempt = 0; attempt < 20; attempt += 1) {
			await settle();
		}
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(viewer().isDirty()).toBe(true);
	});

	/** The Home tab's New Slide command (a button in the ribbon). */
	function newSlideButton(): HTMLButtonElement {
		const button = [...container.querySelectorAll<HTMLButtonElement>('button')].find(
			(candidate) => candidate.getAttribute('aria-label') === 'New Slide',
		);
		if (!button) {
			throw new Error('no New Slide command rendered');
		}
		return button;
	}

	/** Every edit the public handle offers, tried on the active slide. */
	async function tryHandleEdits(viewer: LiveHandle): Promise<void> {
		const shape = { type: 'shape', id: 'probe', x: 0, y: 0, width: 50, height: 50 } as never;
		await act(async () => {
			viewer().addElement(shape);
			viewer().addSlide(0);
			viewer().updateElement('title-0', { x: 500 } as never);
			viewer().deleteElements(['title-0']);
			await Promise.resolve();
		});
		await settle();
	}

	it('opens a file the service leaves read-only: no edit through the handle or the ribbon', async () => {
		const viewer = await mount({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: vi.fn() }, false, 3));
		expect(viewer().getSlideCount()).toBe(3);
		expect(viewer().getMode()).toBe('preview');
		const before = viewer().getElements(0);
		expect(before).toHaveLength(1);
		await act(async () => {
			expect(viewer().addElement({ type: 'shape', id: 'p', x: 0, y: 0 } as never)).toBeUndefined();
		});
		await tryHandleEdits(viewer);
		expect(viewer().getSlideCount()).toBe(3);
		expect(viewer().getElements(0)).toStrictEqual(before);
		expect(viewer().isDirty()).toBe(false);
		// The ribbon: New Slide is disabled, and clicking it changes nothing.
		expect(newSlideButton().disabled).toBe(true);
		click(newSlideButton());
		await settle();
		expect(viewer().getSlideCount()).toBe(3);
		expect(viewer().isDirty()).toBe(false);
	});

	it('edits again when a later cloud file is editable, and when the device deck is opened', async () => {
		const viewer = await mount({ save: vi.fn(async () => undefined) });
		await openFromCloud(await cloudFile({ save: vi.fn() }, false, 3));
		await tryHandleEdits(viewer);
		expect(viewer().getSlideCount()).toBe(3);

		await openFromCloud(await cloudFile({ save: vi.fn() }, true, 4));
		expect(viewer().getSlideCount()).toBe(4);
		expect(newSlideButton().disabled).toBe(false);
		click(newSlideButton());
		await settle();
		expect(viewer().getSlideCount()).toBe(5);
		await tryHandleEdits(viewer);
		expect(viewer().getSlideCount()).toBe(6);
		expect(viewer().getMode()).toBe('edit');
		expect(viewer().isDirty()).toBe(true);
	});
});
