// @vitest-environment happy-dom
/**
 * Host-owned workflows of `useViewerBuildingBlocks` (#285): operations a custom shell calls from its
 * own pickers and dialogs, the requests that raise them, and the permission and history rules that
 * hold while they run.
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
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	data.slides[0].elements = [
		{ id: 'title', type: 'text', x: 67, y: 40, width: 400, height: 60, text: 'Alpha beta alpha' },
	];
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

async function select(id: string) {
	const press = {
		button: 0,
		shiftKey: false,
		metaKey: false,
		ctrlKey: false,
		clientX: 100,
		clientY: 60,
		stopPropagation: vi.fn(),
		preventDefault: vi.fn(),
	} as unknown as React.MouseEvent;
	await act(async () => {
		latest.canvasProps.onMouseDown(id, press);
		latest.canvasProps.onClick(id, press);
	});
}

afterEach(async () => {
	await act(async () => root?.unmount());
	container?.remove();
	onDirtyChange.mockClear();
	onRequest.mockClear();
	vi.unstubAllGlobals();
});

test('the Picture, Video, SmartArt and Equation controls raise host requests, not dead clicks', async () => {
	await mount();
	await act(async () => latest.toolbarProps.onOpenImagePicker?.());
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'image' });
	expect(latest.actions.pending).toStrictEqual({ kind: 'image' });
	await act(async () => latest.toolbarProps.onOpenMediaPicker?.());
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'media' });
	await act(async () => latest.toolbarProps.onAddSmartArt?.());
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'smartArt' });
	await act(async () => latest.toolbarProps.onAddEquation?.());
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'equation', existingOmml: null });
	await act(async () => latest.actions.cancel());
	expect(latest.actions.pending).toBeNull();
});

test('requests are not raised while the viewer is read-only', async () => {
	await mount({ canEdit: false });
	await act(async () => latest.toolbarProps.onOpenImagePicker?.());
	expect(onRequest).not.toHaveBeenCalled();
	expect(latest.actions.availability.insertImage).toMatchObject({
		available: false,
		reason: 'read-only',
	});
	await expect(
		latest.actions.insertMedia(new File(['x'], 'a.mp3', { type: 'audio/mpeg' })),
	).resolves.toStrictEqual({
		ok: false,
		code: 'read-only',
	});
});

test('image and media insertion validate the file and report why one is refused', async () => {
	await mount();
	await expect(
		latest.actions.insertImage(new File(['x'], 'a.pdf', { type: 'application/pdf' })),
	).resolves.toStrictEqual({
		ok: false,
		code: 'unsupported-type',
	});
	await expect(
		latest.actions.insertImage(new File([], 'a.png', { type: 'image/png' })),
	).resolves.toStrictEqual({
		ok: false,
		code: 'empty-file',
	});
	await expect(
		latest.actions.insertMedia(new File(['x'], 'a.png', { type: 'image/png' })),
	).resolves.toStrictEqual({
		ok: false,
		code: 'unsupported-type',
	});
	await mount({ fileLimits: { mediaBytes: 1 } });
	await expect(
		latest.actions.insertMedia(new File(['xx'], 'a.mp3', { type: 'audio/mpeg' })),
	).resolves.toStrictEqual({
		ok: false,
		code: 'file-too-large',
	});
	expect(onDirtyChange).not.toHaveBeenCalledWith(true);
});

test('an audio file is inserted as a selected media element and the deck becomes dirty', async () => {
	await mount();
	let result: unknown;
	await act(async () => {
		result = await latest.actions.insertMedia(new File(['x'], 'a.mp3', { type: 'audio/mpeg' }));
	});
	expect(result).toStrictEqual({ ok: true });
	expect(latest.canvasProps.activeSlide?.elements.some((e) => e.type === 'media')).toBeTruthy();
	expect(onDirtyChange).toHaveBeenCalledWith(true);
});

test('hyperlink: current value, apply, unsafe refusal and remove on the selected element', async () => {
	await mount();
	expect(latest.actions.availability.hyperlink.reason).toBe('no-selection');
	expect(
		latest.actions.hyperlink.apply({
			targetType: 'url',
			url: 'https://a.test',
			tooltip: '',
			emailAddress: '',
			emailSubject: '',
			slideNumber: 1,
			filePath: '',
			actionVerb: 'nextSlide',
		}),
	).toStrictEqual({ ok: false, code: 'no-selection' });
	await select(latest.canvasProps.activeSlide!.elements[0]!.id);
	const data = (url: string) => ({
		targetType: 'url' as const,
		url,
		tooltip: 'Go',
		emailAddress: '',
		emailSubject: '',
		slideNumber: 1,
		filePath: '',
		actionVerb: 'nextSlide' as const,
	});
	expect(latest.actions.hyperlink.apply(data('javascript:alert(1)'))).toStrictEqual({
		ok: false,
		code: 'unsafe-url',
	});
	expect(latest.actions.hyperlink.apply(data(''))).toStrictEqual({
		ok: false,
		code: 'invalid-target',
	});
	await act(async () => {
		expect(latest.actions.hyperlink.apply(data('https://example.test'))).toStrictEqual({
			ok: true,
		});
	});
	expect(latest.actions.hyperlink.current).toMatchObject({
		targetType: 'url',
		url: 'https://example.test',
		tooltip: 'Go',
	});
	expect(onDirtyChange).toHaveBeenCalledWith(true);
	await act(async () => {
		expect(latest.actions.hyperlink.remove()).toStrictEqual({ ok: true });
	});
	expect(latest.actions.hyperlink.current).toBeNull();
	expect(latest.actions.hyperlink.remove()).toStrictEqual({ ok: false, code: 'no-hyperlink' });
});

test('SmartArt and equation inserts add a selected element; the dialog requests clear on completion', async () => {
	await mount();
	await act(async () => latest.toolbarProps.onAddSmartArt?.());
	expect(latest.actions.pending).toStrictEqual({ kind: 'smartArt' });
	await act(async () => {
		expect(latest.actions.smartArt.insert('basic-block-list' as never, ['A', 'B'])).toStrictEqual({
			ok: true,
		});
	});
	expect(latest.actions.pending).toBeNull();
	expect(latest.canvasProps.activeSlide?.elements.some((e) => e.type === 'smartArt')).toBeTruthy();
	await act(async () => {
		expect(latest.actions.equation.insert({ 'm:oMath': {} })).toStrictEqual({ ok: true });
	});
	expect(latest.actions.equation.current).toStrictEqual({ 'm:oMath': {} });
	await act(async () => {
		expect(latest.actions.equation.update({ 'm:oMath': { x: 1 } })).toStrictEqual({ ok: true });
	});
	expect(latest.actions.equation.current).toStrictEqual({ 'm:oMath': { x: 1 } });
});

test('the Insert Chart control raises a request and chart.insert adds the chosen chart with data', async () => {
	await mount();
	expect(latest.actions.availability.insertChart).toStrictEqual({
		available: true,
		needsHostUi: true,
	});
	await act(async () => latest.toolbarProps.onOpenChartDialog?.());
	expect(onRequest).toHaveBeenLastCalledWith({ kind: 'chart' });
	expect(latest.actions.pending).toStrictEqual({ kind: 'chart' });
	const before = latest.canvasProps.activeSlide?.elements.length ?? 0;
	expect(
		latest.actions.chart.insert('pie', {
			data: { categories: ['A', 'B'], series: [{ name: 'S', values: [1] }] },
		}),
	).toStrictEqual({ ok: false, code: 'invalid-data' });
	await act(async () => {
		expect(
			latest.actions.chart.insert('pie', {
				title: 'Mix',
				data: { categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] },
			}),
		).toStrictEqual({ ok: true });
	});
	expect(latest.actions.pending).toBeNull();
	const elements = latest.canvasProps.activeSlide?.elements ?? [];
	expect(elements).toHaveLength(before + 1);
	expect(elements.at(-1)).toMatchObject({
		type: 'chart',
		chartData: { chartType: 'pie', title: 'Mix', categories: ['A', 'B'] },
	});
	expect(onDirtyChange).toHaveBeenCalledWith(true);
});

test('chart.insert is refused in a view-only viewer', async () => {
	await mount({ canEdit: false });
	expect(latest.actions.chart.insert('column')).toStrictEqual({ ok: false, code: 'read-only' });
});

test('find and replace expose options, matches, the current match, navigation and both replaces', async () => {
	await mount();
	expect(latest.actions.findReplace.replaceAll()).toStrictEqual({ ok: false, code: 'no-match' });
	await act(async () => latest.actions.findReplace.setQuery('alpha'));
	expect(latest.actions.findReplace.matches).toHaveLength(2);
	expect(latest.actions.findReplace.currentIndex).toBe(0);
	expect(latest.actions.findReplace.currentMatch).not.toBeNull();
	await act(async () => latest.actions.findReplace.next());
	expect(latest.actions.findReplace.currentIndex).toBe(1);
	await act(async () => latest.actions.findReplace.previous());
	expect(latest.actions.findReplace.currentIndex).toBe(0);
	await act(async () => latest.actions.findReplace.setOptions({ matchCase: true }));
	expect(latest.actions.findReplace.options.matchCase).toBeTruthy();
	expect(latest.actions.findReplace.matches).toHaveLength(1);
	await act(async () => latest.actions.findReplace.setOptions({ matchCase: false }));
	await act(async () => latest.actions.findReplace.setReplacement('gamma'));
	await act(async () => {
		expect(latest.actions.findReplace.replaceAll()).toStrictEqual({ ok: true });
	});
	expect(latest.canvasProps.activeSlide?.elements[0]).toMatchObject({ text: 'gamma beta gamma' });
	expect(onDirtyChange).toHaveBeenCalledWith(true);
});

test('replace is refused in a view-only viewer while searching still works', async () => {
	await mount({ canEdit: false });
	await act(async () => latest.actions.findReplace.setQuery('alpha'));
	expect(latest.actions.findReplace.matches).toHaveLength(2);
	expect(latest.actions.findReplace.replaceAll()).toStrictEqual({ ok: false, code: 'read-only' });
	expect(latest.actions.findReplace.replaceCurrent()).toStrictEqual({
		ok: false,
		code: 'read-only',
	});
});

test('layout apply and reset are wired, validated against the slide master, and available', async () => {
	await mount();
	expect(latest.toolbarProps.onApplyLayout).toBeTypeOf('function');
	expect(latest.toolbarProps.onResetSlide).toBeTypeOf('function');
	expect(latest.actions.layout.available.length).toBeGreaterThan(0);
	expect(latest.actions.layout.currentPath).toBeTypeOf('string');
	expect(latest.actions.layout.current?.path).toBe(latest.actions.layout.currentPath);
	await expect(latest.actions.layout.apply('ppt/slideLayouts/nope.xml')).resolves.toStrictEqual({
		ok: false,
		code: 'unknown-layout',
	});
	const target = latest.actions.layout.available.find(
		(layout) => layout.path !== latest.actions.layout.currentPath,
	);
	if (target) {
		let result: unknown;
		await act(async () => {
			result = await latest.actions.layout.apply(target.path);
		});
		expect(result).toStrictEqual({ ok: true });
		expect(latest.actions.layout.currentPath).toBe(target.path);
	}
	let reset: unknown;
	await act(async () => {
		reset = await latest.actions.layout.reset();
	});
	expect(reset).toStrictEqual({ ok: true });
});

test('unmounted dialog toolbar actions are listed so a shell can hide or replace them', async () => {
	await mount();
	expect(latest.actions.unmountedDialogActions).toContain('onOpenDocumentProperties');
	expect(latest.actions.unmountedDialogActions).not.toContain('onOpenChartDialog');
	for (const name of latest.actions.unmountedDialogActions) {
		expect(Object.keys(latest.toolbarProps)).toContain(name);
	}
});

test('a video whose size probe finishes after the slide changed is not added to the wrong slide', async () => {
	await mount();
	const probes: Array<{ onloadedmetadata?: () => void; videoWidth: number; videoHeight: number }> =
		[];
	const realCreate = document.createElement.bind(document);
	vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => {
		if (tag !== 'video') {
			return realCreate(tag);
		}
		const probe = { videoWidth: 320, videoHeight: 180 } as (typeof probes)[number];
		probes.push(probe);
		return probe as unknown as HTMLVideoElement;
	}) as typeof document.createElement);
	const firstSlideId = latest.canvasProps.activeSlide!.id;
	const before = latest.canvasProps.activeSlide!.elements.length;
	let result: unknown;
	let pending: Promise<unknown> | undefined;
	await act(async () => {
		pending = latest.actions.insertMedia(new File(['x'], 'a.mp4', { type: 'video/mp4' }));
		await vi.waitFor(() => expect(probes).toHaveLength(1));
	});
	await act(async () => latest.slidesPaneProps.onSelectSlide(1));
	expect(latest.canvasProps.activeSlide!.id).not.toBe(firstSlideId);
	await act(async () => {
		probes[0]!.onloadedmetadata?.();
		result = await pending;
	});
	vi.restoreAllMocks();
	expect(result).toStrictEqual({ ok: false, code: 'stale-target' });
	expect(latest.canvasProps.activeSlide!.elements.some((e) => e.type === 'media')).toBeFalsy();
	await act(async () => latest.slidesPaneProps.onSelectSlide(0));
	expect(latest.canvasProps.activeSlide!.elements).toHaveLength(before);
});
