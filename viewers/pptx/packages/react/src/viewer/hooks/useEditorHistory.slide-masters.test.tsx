// @vitest-environment happy-dom
/**
 * A slide paste commits its slides and, in the same commit, the slide masters the pasted deck
 * brings. The snapshots record each master's theme part, so the masters arriving a render after the
 * slides must not read as a second undo step: one Undo takes the whole paste back. A master added
 * or deleted in Master view is no step either, as before the theme parts were recorded.
 */
import type { PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import React, { act, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';

import type { CanvasSize } from '../types';
import { useEditorHistory } from './useEditorHistory';

function slide(id: string): PptxSlide {
	return { id, rId: id, slideNumber: 1, elements: [] } as unknown as PptxSlide;
}

function master(path: string, themePath: string): PptxSlideMaster {
	return { path, themePath } as unknown as PptxSlideMaster;
}

const MASTERS = [master('ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme1.xml')];
const PASTED_MASTERS = [
	...MASTERS,
	master('ppt/slideMasters/slideMaster2.xml', 'ppt/theme/theme2.xml'),
];

interface HarnessApi {
	slideCount: () => number;
	canUndo: () => boolean;
	paste: (passMasters: boolean) => void;
	setMasters: (masters: PptxSlideMaster[]) => void;
	undo: () => void;
}

let api: HarnessApi | null = null;

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

function Harness(): React.ReactElement {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [slide('a'), slide('b')]);
	const [slideMasters, setSlideMasters] = useState<PptxSlideMaster[]>(MASTERS);
	const [canvasSize, setCanvasSize] = useState<CanvasSize>({ width: 960, height: 540 });
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
		slideMasters,
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId: setTemplates,
		setSelectedElementId: noop<string | null>(),
		setSelectedElementIds: noop<string[]>(),
		setEditTemplateMode: noop<boolean>(),
		setHeaderFooter: noop<never>(),
	});
	const slidesRef = useRef(slides);
	slidesRef.current = slides;
	const historyRef = useRef(history);
	historyRef.current = history;
	api = {
		slideCount: () => slidesRef.current.length,
		canUndo: () => historyRef.current.canUndo,
		paste: (passMasters) => {
			const [first, ...rest] = slidesRef.current;
			const next = [first!, slide('pasted'), ...rest];
			historyRef.current.commitSlides(
				next,
				undefined,
				undefined,
				passMasters ? PASTED_MASTERS : undefined,
			);
			setSlideMasters(PASTED_MASTERS);
			historyRef.current.markDirty();
		},
		setMasters: (masters) => {
			setSlideMasters(masters);
			setSlides((previous) => [...previous]);
			historyRef.current.markDirty();
		},
		undo: () => historyRef.current.handleUndo(),
	};
	return <div />;
}

let root: Root | null = null;
const previousActEnvironment = (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean })
	.IS_REACT_ACT_ENVIRONMENT;

async function mount(): Promise<HarnessApi> {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	const container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	await act(async () => {
		root!.render(<Harness />);
	});
	return api!;
}

afterEach(async () => {
	await act(async () => {
		root?.unmount();
	});
	root = null;
	api = null;
	document.body.replaceChildren();
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
		previousActEnvironment;
});

describe('useEditorHistory slide paste with its own masters', () => {
	it('takes the pasted slides back in one Undo', async () => {
		const harness = await mount();
		await act(async () => harness.paste(true));
		expect(harness.slideCount()).toBe(3);
		await act(async () => harness.undo());
		expect(harness.slideCount()).toBe(2);
		expect(harness.canUndo()).toBe(false);
	});

	it('takes the pasted slides back in one Undo when the masters land unannounced', async () => {
		const harness = await mount();
		await act(async () => harness.paste(false));
		await act(async () => harness.undo());
		expect(harness.slideCount()).toBe(2);
		expect(harness.canUndo()).toBe(false);
	});

	it('records no step for a master added or deleted in Master view', async () => {
		const harness = await mount();
		await act(async () => harness.setMasters(PASTED_MASTERS));
		expect(harness.canUndo()).toBe(false);
		await act(async () => harness.setMasters(MASTERS));
		expect(harness.canUndo()).toBe(false);
	});

	it('still records a master moved onto another theme part', async () => {
		const harness = await mount();
		await act(async () =>
			harness.setMasters([master('ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme2.xml')]),
		);
		expect(harness.canUndo()).toBe(true);
	});
});
