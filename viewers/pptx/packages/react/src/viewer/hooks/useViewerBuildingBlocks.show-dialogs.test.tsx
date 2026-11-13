// @vitest-environment happy-dom
/**
 * Host-owned slide show dialogs of `useViewerBuildingBlocks`: Set Up Slide Show, Custom Shows,
 * Broadcast (unsupported) and the shortcut help, their requests, permission rechecks and dirty
 * notifications.
 */
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, expect, test, vi } from 'vitest';

import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type {
	BuildingBlockRequest,
	UseViewerBuildingBlocksInput,
	ViewerBuildingBlocksResult,
} from './useViewerBuildingBlocks';

let root: Root;
let container: HTMLDivElement;
let latest: ViewerBuildingBlocksResult;
(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;
const onDirtyChange = vi.fn();
const onRequest = vi.fn<(request: BuildingBlockRequest) => void>();

function Harness({
	content,
	options,
}: {
	content: Uint8Array;
	options: Partial<UseViewerBuildingBlocksInput>;
}) {
	latest = useViewerBuildingBlocks({
		content,
		canEdit: true,
		autosaveEnabled: false,
		onDirtyChange,
		onRequest,
		...options,
	});
	return <div />;
}

async function mount(options: Partial<UseViewerBuildingBlocksInput> = {}) {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 3 });
	data.slides[0].isDirty = true;
	const content = await handler.save(data.slides);
	handler.dispose();
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	await act(async () => root.render(<Harness content={content} options={options} />));
	const deadline = Date.now() + 10000;
	while (latest.loading && Date.now() < deadline) {
		await act(async () => {
			await new Promise<void>((resolve) => {
				setTimeout(resolve, 0);
			});
		});
	}
	expect(latest.error).toBeNull();
}

afterEach(async () => {
	await act(async () => root?.unmount());
	container?.remove();
	onDirtyChange.mockClear();
	onRequest.mockClear();
});

test('Set Up Slide Show raises a request, applies a patch and marks the deck dirty', async () => {
	await mount();
	expect(latest.actions.availability.setUpSlideShow).toEqual({
		available: true,
		needsHostUi: true,
	});
	await act(async () => latest.toolbarProps.onOpenSetUpSlideShow?.());
	expect(onRequest).toHaveBeenCalledWith({ kind: 'setUpSlideShow' });
	expect(latest.actions.pending).toEqual({ kind: 'setUpSlideShow' });
	expect(latest.actions.setUpSlideShow.current.showType).toBe('presented');
	onDirtyChange.mockClear();

	let result: unknown;
	await act(async () => {
		result = latest.actions.setUpSlideShow.apply({ showType: 'kiosk' });
	});
	expect(result).toEqual({ ok: true });
	expect(latest.actions.setUpSlideShow.current).toMatchObject({ showType: 'kiosk', loop: true });
	expect(latest.actions.pending).toBeNull();
	expect(onDirtyChange).toHaveBeenCalledWith(true);

	await act(async () => {
		result = latest.actions.setUpSlideShow.apply({ showSlidesMode: 'range', showSlidesTo: 99 });
	});
	expect(result).toEqual({ ok: false, code: 'invalid-range' });
});

test('Set Up Slide Show is refused when the viewer is read-only', async () => {
	await mount({ canEdit: false });
	expect(latest.actions.availability.setUpSlideShow.available).toBe(false);
	await act(async () => latest.toolbarProps.onOpenSetUpSlideShow?.());
	expect(onRequest).not.toHaveBeenCalled();
	expect(latest.actions.setUpSlideShow.apply({ loopContinuously: true })).toEqual({
		ok: false,
		code: 'read-only',
	});
});

test('Custom Shows lists, creates, renames, edits, deletes and starts a show', async () => {
	await mount();
	await act(async () => latest.toolbarProps.onOpenCustomShows?.());
	expect(onRequest).toHaveBeenCalledWith({ kind: 'customShows' });
	const rIds = latest.actions.customShows.slides.map((slide) => slide.rId);
	expect(rIds).toHaveLength(3);

	await act(async () => {
		expect(latest.actions.customShows.create('Short', [rIds[0] as string])).toEqual({ ok: true });
	});
	const created = latest.actions.customShows.shows[0];
	expect(created).toMatchObject({ name: 'Short', slideCount: 1 });
	expect(onDirtyChange).toHaveBeenCalledWith(true);
	const id = created?.id as string;

	await act(async () => {
		latest.actions.customShows.rename(id, 'Intro');
		latest.actions.customShows.setSlides(id, [rIds[2] as string, rIds[1] as string]);
	});
	expect(latest.actions.customShows.shows[0]).toMatchObject({
		name: 'Intro',
		slideRIds: [rIds[2], rIds[1]],
	});
	expect(latest.actions.customShows.setSlides(id, ['nope'])).toEqual({
		ok: false,
		code: 'unknown-slide',
	});

	await act(async () => {
		expect(latest.actions.customShows.start(id)).toEqual({ ok: true });
	});
	expect(latest.actions.customShows.activeId).toBe(id);
	expect(latest.mode).toBe('present');

	// While presenting the viewer is not in edit mode, so the edit gate refuses at execution.
	expect(latest.actions.customShows.remove(id)).toEqual({ ok: false, code: 'read-only' });
	expect(latest.actions.customShows.shows).toHaveLength(1);
});

test('Broadcast is reported unsupported and raises no request', async () => {
	await mount();
	expect(latest.actions.availability.broadcast).toEqual({
		available: false,
		reason: 'unsupported',
		needsHostUi: true,
	});
	expect(latest.actions.broadcast.supported).toBe(false);
	expect(latest.actions.unmountedDialogActions).toContain('onOpenBroadcastDialog');
	expect(latest.actions.unmountedDialogActions).not.toContain('onOpenSetUpSlideShow');
	await act(async () => latest.toolbarProps.onOpenBroadcastDialog?.());
	expect(onRequest).not.toHaveBeenCalled();
});

test('the shortcut help toggles, emits events and exposes the catalogue', async () => {
	await mount();
	expect(latest.actions.shortcuts.catalogue.length).toBeGreaterThan(20);
	expect(latest.actions.shortcuts.catalogue[0]).toMatchObject({ id: expect.any(String) });
	await act(async () => latest.toolbarProps.onToggleShortcuts());
	expect(latest.actions.shortcuts.open).toBe(true);
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'shortcuts', open: true });
	expect(latest.actions.pending).toEqual({ kind: 'shortcuts', open: true });
	await act(async () => latest.actions.shortcuts.toggle());
	expect(latest.actions.shortcuts.open).toBe(false);
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'shortcuts', open: false });
	expect(latest.actions.pending).toBeNull();
});
