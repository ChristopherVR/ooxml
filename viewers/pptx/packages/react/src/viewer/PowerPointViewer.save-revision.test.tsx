// @vitest-environment happy-dom
/**
 * Revision-aware acknowledgement of a Save the host persists itself, end to end through the real
 * `PowerPointViewer`: a snapshot leaves the deck dirty, only the current revision cleans it, and an
 * edit made while the host was persisting stays unsaved.
 */
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

async function mount(): Promise<LiveHandle> {
	const bytes = await sampleDeck();
	const ref = createRef<ViewerHandle>();
	await act(async () => {
		root.render(<PowerPointViewer ref={ref} content={bytes} canEdit />);
	});
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0) {
			return () => ref.current!;
		}
		await settle();
	}
	throw new Error('the viewer never finished loading its deck');
}

async function edit(viewer: LiveHandle): Promise<void> {
	await act(async () => {
		viewer().addSlide(0);
		await Promise.resolve();
	});
}

describe('PowerPointViewer host-persisted save acknowledgement', () => {
	it('stays dirty after a snapshot and reads clean once its revision is acknowledged', async () => {
		const viewer = await mount();
		await edit(viewer);
		expect(viewer().isDirty()).toBe(true);
		const snapshot = await viewer().getSaveSnapshot();
		expect(snapshot.content.byteLength).toBeGreaterThan(0);
		expect(viewer().isDirty()).toBe(true);
		let result: ReturnType<ViewerHandle['acknowledgeSave']> | undefined;
		await act(async () => {
			result = viewer().acknowledgeSave(snapshot.revision);
		});
		expect(result?.status).toBe('clean');
		expect(viewer().isDirty()).toBe(false);
		expect(viewer().canUndo()).toBe(true);
	});

	it('keeps an edit made during persistence dirty and ignores a stale acknowledgement', async () => {
		const viewer = await mount();
		await edit(viewer);
		const first = await viewer().getSaveSnapshot();
		await edit(viewer);
		let pending: ReturnType<ViewerHandle['acknowledgeSave']> | undefined;
		await act(async () => {
			pending = viewer().acknowledgeSave(first.revision);
		});
		expect(pending?.status).toBe('pending');
		expect(viewer().isDirty()).toBe(true);
		const second = await viewer().getSaveSnapshot();
		await act(async () => {
			viewer().acknowledgeSave(second.revision);
		});
		expect(viewer().isDirty()).toBe(false);
		await edit(viewer);
		let stale: ReturnType<ViewerHandle['acknowledgeSave']> | undefined;
		await act(async () => {
			stale = viewer().acknowledgeSave(first.revision);
		});
		expect(stale?.status).toBe('stale');
		expect(viewer().isDirty()).toBe(true);
	});
});
