// @vitest-environment happy-dom
/**
 * File > Save through host-owned storage, end to end through the real `PowerPointViewer`: the
 * document reads clean only once `storage.save` has resolved, stays unsaved when it rejects, and
 * stays unsaved when an edit lands while the upload is in flight.
 */
import type { DocumentStorage } from 'ooxml-ui/pptx';
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

async function mount(storage: DocumentStorage): Promise<LiveHandle> {
	const bytes = await sampleDeck();
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(<PowerPointViewer ref={ref} content={bytes} canEdit storage={storage} />);
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

async function editAndSave(viewer: LiveHandle): Promise<void> {
	// Inserted between two slides, so saving renumbers the one after it.
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
}

describe('PowerPointViewer dirty state after a save to host storage', () => {
	it('reads clean once storage.save has resolved', async () => {
		const save = vi.fn(async () => undefined);
		const viewer = await mount({ save });
		await editAndSave(viewer);
		expect(save).toHaveBeenCalledOnce();
		expect(viewer().isDirty()).toBe(false);
		// Saving wrote the saved slide identity back onto the slides; leaving the new slide is no edit.
		await act(async () => {
			viewer().setActiveSlideIndex(0);
		});
		await settle();
		expect(viewer().isDirty()).toBe(false);
	});

	it('stays unsaved when storage.save rejects', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const save = vi.fn(() => Promise.reject(new Error('offline')));
		const viewer = await mount({ save });
		await editAndSave(viewer);
		expect(save).toHaveBeenCalledOnce();
		expect(viewer().isDirty()).toBe(true);
	});

	it('stays unsaved when an edit lands while the upload is in flight', async () => {
		let live: LiveHandle | undefined;
		const save = vi.fn(async () => {
			live?.().addSlide();
			await new Promise((resolve) => {
				setTimeout(resolve, 20);
			});
		});
		const viewer = await mount({ save });
		live = viewer;
		await editAndSave(viewer);
		expect(save).toHaveBeenCalledOnce();
		expect(viewer().isDirty()).toBe(true);
	});
});
