// @vitest-environment happy-dom
/**
 * A Slide Master view command replaces the core handler, but it is an edit, not a
 * load: the assistant's staged proposals must stay listed and still apply. Only a
 * new presentation load (the bridge's deck load) drops them.
 */
import type { MasterViewTarget } from 'ooxml-ui/pptx';
import { ProposalStore } from 'ooxml-ui/pptx/ai';
import type { PptxAiBridge } from 'ooxml-ui/pptx/ai';
import type { PptxHandler, PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import { PptxHandler as PptxHandlerCtor } from 'pptx-viewer-core';
import React, { act, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useAiBridge } from './ai/useAiBridge';
import { useEditorHistory } from './useEditorHistory';
import type { UseMasterViewCrudResult } from './useMasterViewCrud';
import { useMasterViewCrud } from './useMasterViewCrud';

interface Handle {
	bridge: PptxAiBridge;
	crud: UseMasterViewCrudResult;
	slides: PptxSlide[];
	handler: () => PptxHandler | null;
}

const noop = (): void => {};
const TARGET: MasterViewTarget = { tab: 'slides', masterIndex: 0, layoutIndex: null };

function Harness({
	handler,
	initial,
	deckLoad,
	onReady,
}: {
	handler: PptxHandler;
	initial: { slides: PptxSlide[]; slideMasters: PptxSlideMaster[] };
	deckLoad: number;
	onReady: (handle: Handle) => void;
}): null {
	const handlerRef = useRef<PptxHandler | null>(handler);
	const [slides, setSlides] = useState(initial.slides);
	const [slideMasters, setSlideMasters] = useState(initial.slideMasters);
	const [canvasSize, setCanvasSize] = useState({ width: 960, height: 540 });
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [templates, setTemplates] = useState({});
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
		setSelectedElementId: noop,
		setSelectedElementIds: noop,
		setEditTemplateMode: noop,
		setHeaderFooter: noop,
	});
	const crud = useMasterViewCrud({
		handlerRef,
		slides,
		slideMasters,
		target: TARGET,
		setSlideMasters,
		setLayoutOptions: noop,
		setActiveMasterIndex: noop,
		setActiveLayoutIndex: noop,
		markDirty: history.markDirty,
		forgetHistory: history.forgetHistory,
		pushToast: vi.fn(),
	});
	const bridge = useAiBridge({
		slides,
		activeSlideIndex,
		canvasSize,
		theme: undefined,
		selectedElementId: null,
		selectedElementIds: [],
		pinnedFocus: null,
		handlerRef,
		deckLoad,
		sections: [],
		customShows: [],
		presentationProperties: {},
		customProperties: [],
		coreProperties: undefined,
		appProperties: undefined,
		viewProperties: undefined,
		tableStyleMap: undefined,
		tableStylesDefaultId: undefined,
		tagCollections: [],
		setCanvasSize,
		setSections: noop,
		setCustomShows: noop,
		setPresentationProperties: noop,
		setCustomProperties: noop,
		setCoreProperties: noop,
		setAppProperties: noop,
		setViewProperties: noop,
		setTableStyleMap: noop,
		setTableStylesDefaultId: noop,
		setTagCollections: noop,
		setSlides,
		setActiveSlideIndex,
		applySelection: noop,
		recordSlidesStep: history.recordSlidesCommit,
		markDirty: history.markDirty,
		applyThemeUpdates: noop,
	});
	onReady({ bridge, crud, slides, handler: () => handlerRef.current });
	return null;
}

let container: HTMLDivElement;
let root: Root;
let handle: Handle;
let handler: PptxHandler;
let initial: { slides: PptxSlide[]; slideMasters: PptxSlideMaster[] };

function render(deckLoad: number): void {
	act(() => {
		root.render(
			<Harness
				handler={handler}
				initial={initial}
				deckLoad={deckLoad}
				onReady={(next) => {
					handle = next;
				}}
			/>,
		);
	});
}

function stageRedBackground(store: ProposalStore): void {
	store.stage('Red background', (slides) =>
		slides.map((slide, index) => (index === 0 ? { ...slide, backgroundColor: '#FF0000' } : slide)),
	);
}

beforeEach(async () => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const created = await PptxHandlerCtor.create({ initialSlideCount: 1 });
	handler = created.handler;
	initial = { slides: created.data.slides, slideMasters: created.data.slideMasters ?? [] };
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	render(1);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	handler.dispose();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('AI proposals across Slide Master view commands', () => {
	it('keeps a staged proposal listed and applicable after Insert Layout replaces the handler', async () => {
		const store = new ProposalStore(handle.bridge);
		stageRedBackground(store);

		await act(async () => {
			await handle.crud.handleCrudAction('addLayout');
		});

		expect(handle.handler()).not.toBe(handler);
		expect(store.list().map((view) => view.label)).toStrictEqual(['Red background']);
		expect(store.discarded).toBeFalsy();
		let applied = 0;
		act(() => {
			applied = store.acceptAll();
		});
		expect(applied).toBe(1);
		expect(handle.slides[0]?.backgroundColor).toBe('#FF0000');
	});

	it('drops the staged proposal when another presentation loads', () => {
		const store = new ProposalStore(handle.bridge);
		stageRedBackground(store);

		render(2);

		expect(store.list()).toStrictEqual([]);
		expect(store.discarded).toBeTruthy();
	});
});
