import type { MasterViewTarget } from 'ooxml-ui/pptx';
// @vitest-environment happy-dom
import type { PptxHandler, PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import { PptxHandler as PptxHandlerCtor } from 'pptx-viewer-core';
import React, { act, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { CanvasSize } from '../types';
import type { EditorHistoryResult } from './useEditorHistory';
import { useEditorHistory } from './useEditorHistory';
import type { UseMasterViewCrudResult } from './useMasterViewCrud';
import { useMasterViewCrud } from './useMasterViewCrud';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

interface HarnessHandle {
	handleCrudAction: (id: 'addLayout' | 'deleteLayout') => Promise<void>;
	getCrudActions: () => ReturnType<typeof useMasterViewCrud>['crudActions'];
	getSlideMasters: () => PptxSlideMaster[];
	getLayoutOptions: () => readonly { path: string }[];
}

function Harness({
	handler,
	initialData,
	target,
	onReady,
}: {
	handler: PptxHandler;
	initialData: {
		slides: PptxSlide[];
		slideMasters: PptxSlideMaster[];
		layoutOptions?: { path: string; name: string }[];
	};
	target: MasterViewTarget;
	onReady: (handle: HarnessHandle) => void;
}): React.ReactElement {
	const handlerRef = useRef(handler);
	// A command leaves the slides alone (nothing sets them), so they are the props' own array.
	const slides = initialData.slides;
	const [slideMasters, setSlideMasters] = useState(initialData.slideMasters);
	const [layoutOptions, setLayoutOptions] = useState(initialData.layoutOptions ?? []);
	const [, setActiveMasterIndex] = useState(target.masterIndex);
	const [, setActiveLayoutIndex] = useState(target.layoutIndex);
	const markDirty = vi.fn();
	const pushToast = vi.fn();

	const result = useMasterViewCrud({
		handlerRef,
		slides,
		slideMasters,
		target,
		setSlideMasters,
		setLayoutOptions,
		setActiveMasterIndex,
		setActiveLayoutIndex,
		markDirty,
		forgetHistory: vi.fn(),
		pushToast,
	});

	onReady({
		handleCrudAction: result.handleCrudAction,
		getCrudActions: () => result.crudActions,
		getSlideMasters: () => slideMasters,
		getLayoutOptions: () => layoutOptions,
	});

	return <div />;
}

interface HistoryHarnessHandle {
	slides: PptxSlide[];
	slideMasters: PptxSlideMaster[];
	history: EditorHistoryResult;
	crud: UseMasterViewCrudResult;
	setTarget: (target: MasterViewTarget) => void;
}

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

/** The real editor history and the CRUD hook, wired the way `PowerPointViewer` wires them. */
function HistoryHarness({
	handler,
	initialData,
	onReady,
}: {
	handler: PptxHandler;
	initialData: { slides: PptxSlide[]; slideMasters: PptxSlideMaster[] };
	onReady: (handle: HistoryHarnessHandle) => void;
}): React.ReactElement {
	const handlerRef = useRef(handler);
	const [slides, setSlides] = useState(initialData.slides);
	const [slideMasters, setSlideMasters] = useState(initialData.slideMasters);
	const [canvasSize, setCanvasSize] = useState<CanvasSize>({ width: 960, height: 540 });
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [templates, setTemplates] = useState({});
	const [target, setTarget] = useState<MasterViewTarget>({
		tab: 'slides',
		masterIndex: 0,
		layoutIndex: null,
	});
	const history = useEditorHistory({
		slides,
		canvasSize,
		activeSlideIndex,
		templateElementsBySlideId: templates,
		selectedElementId: null,
		selectedElementIds: [],
		editTemplateMode: false,
		headerFooter: {} as never,
		loading: false,
		error: null,
		hasActivePointerInteraction: useCallback(() => false, []),
		pointerCommitNonce: 0,
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId: setTemplates,
		setSelectedElementId: noop<string | null>(),
		setSelectedElementIds: noop<string[]>(),
		setEditTemplateMode: noop<boolean>(),
		setHeaderFooter: noop<never>(),
	});
	const crud = useMasterViewCrud({
		handlerRef,
		slides,
		slideMasters,
		target,
		setSlideMasters,
		setLayoutOptions: noop<{ path: string; name: string }[]>(),
		setActiveMasterIndex: noop<number>(),
		setActiveLayoutIndex: noop<number | null>(),
		markDirty: history.markDirty,
		forgetHistory: history.forgetHistory,
		pushToast: vi.fn(),
	});
	onReady({ slides, slideMasters, history, crud, setTarget });
	return <div />;
}

describe('useMasterViewCrud', () => {
	it('disables deleteLayout for a layout still used by a slide', async () => {
		const { handler, data } = await PptxHandlerCtor.create({ initialSlideCount: 1 });
		try {
			const master = data.slideMasters![0];
			const inUseLayout = master.layouts!.find((l) => l.path === data.slides[0].layoutPath)!;
			const target: MasterViewTarget = {
				tab: 'slides',
				masterIndex: 0,
				layoutIndex: master.layouts!.indexOf(inUseLayout),
			};

			let handle: HarnessHandle | null = null;
			act(() => {
				root.render(
					<Harness
						handler={handler}
						initialData={{ slides: data.slides, slideMasters: data.slideMasters! }}
						target={target}
						onReady={(h) => {
							handle = h;
						}}
					/>,
				);
			});

			const deleteAction = handle!.getCrudActions().find((a) => a.id === 'deleteLayout');
			expect(deleteAction?.enabled).toBeFalsy();
			expect(deleteAction?.disabledReasonKey).toBe('pptx.masterView.layoutInUse');
		} finally {
			handler.dispose();
		}
	});

	it('addLayout adopts the returned handler/data: a new layout appears under the master', async () => {
		const { handler, data } = await PptxHandlerCtor.create({ initialSlideCount: 1 });
		try {
			const target: MasterViewTarget = { tab: 'slides', masterIndex: 0, layoutIndex: null };
			const before = data.slideMasters![0].layouts!.length;

			let handle: HarnessHandle | null = null;
			act(() => {
				root.render(
					<Harness
						handler={handler}
						initialData={{ slides: data.slides, slideMasters: data.slideMasters! }}
						target={target}
						onReady={(h) => {
							handle = h;
						}}
					/>,
				);
			});

			await act(async () => {
				await handle!.handleCrudAction('addLayout');
			});

			expect(handle!.getSlideMasters()[0].layouts).toHaveLength(before + 1);
		} finally {
			handler.dispose();
		}
	});

	it('adopts the layout list a command returns, so Home > Layout offers what the deck has', async () => {
		const { handler, data } = await PptxHandlerCtor.create({ initialSlideCount: 1 });
		try {
			const target: MasterViewTarget = { tab: 'slides', masterIndex: 0, layoutIndex: null };
			const loaded = data.layoutOptions ?? [];

			let handle: HarnessHandle | null = null;
			act(() => {
				root.render(
					<Harness
						handler={handler}
						initialData={{
							slides: data.slides,
							slideMasters: data.slideMasters!,
							layoutOptions: loaded,
						}}
						target={target}
						onReady={(h) => {
							handle = h;
						}}
					/>,
				);
			});

			await act(async () => {
				await handle!.handleCrudAction('addLayout');
			});

			const layoutPaths = handle!
				.getSlideMasters()
				.flatMap((master) => (master.layouts ?? []).map((layout) => layout.path));
			expect(handle!.getLayoutOptions().map((option) => option.path)).toStrictEqual(layoutPaths);
			expect(handle!.getLayoutOptions()).toHaveLength(loaded.length + 1);
		} finally {
			handler.dispose();
		}
	});

	it('leaves no undo or redo step that puts a slide back on a deleted layout', async () => {
		const { handler, data } = await PptxHandlerCtor.create({ initialSlideCount: 1 });
		try {
			const layouts = data.slideMasters![0].layouts!;
			const home = data.slides[0].layoutPath!;
			const spare = layouts.find((layout) => layout.path !== home)!;
			let handle: HistoryHarnessHandle | null = null;
			act(() => {
				root.render(
					<HistoryHarness
						handler={handler}
						initialData={{ slides: data.slides, slideMasters: data.slideMasters! }}
						onReady={(h) => {
							handle = h;
						}}
					/>,
				);
			});
			const state = () => handle!;
			const edit = (label: string, patch: Partial<PptxSlide>) =>
				act(() => {
					state().history.commitSlides([{ ...state().slides[0], ...patch }], label);
				});
			const layoutsInUse = () => state().slides.map((slide) => slide.layoutPath);

			// The spare layout is in an undo step (before 'two') and, once 'four' is undone, in the
			// redo step, but no slide uses it now.
			edit('one', { layoutPath: spare.path });
			edit('two', { layoutPath: home });
			edit('three', { backgroundColor: '#FF0000' });
			edit('four', { layoutPath: spare.path });
			act(() => state().history.handleUndo());
			expect(layoutsInUse()).toStrictEqual([home]);
			act(() =>
				state().setTarget({ tab: 'slides', masterIndex: 0, layoutIndex: layouts.indexOf(spare) }),
			);

			await act(async () => {
				await state().crud.handleCrudAction('deleteLayout');
			});

			expect(state().slideMasters[0].layouts!.map((layout) => layout.path)).not.toContain(
				spare.path,
			);
			expect(state().history).toMatchObject({ canUndo: true, canRedo: false });
			act(() => state().history.handleUndo());
			expect(layoutsInUse()).toStrictEqual([home]);
			expect(state().slides[0].backgroundColor).not.toBe('#FF0000');
			expect(state().history).toMatchObject({ canUndo: false, canRedo: true });
			act(() => state().history.handleRedo());
			expect(layoutsInUse()).toStrictEqual([home]);
			expect(state().history).toMatchObject({ canUndo: true, canRedo: false });
		} finally {
			handler.dispose();
		}
	});
});
