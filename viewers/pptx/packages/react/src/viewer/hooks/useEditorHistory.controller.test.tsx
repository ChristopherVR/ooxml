// @vitest-environment happy-dom
/**
 * The React history keeps its change detection but no longer its own stack: the
 * real {@link useEditorHistory} records into the shared history controller, so
 * its default depth is the one every binding uses (React kept 120 before) and
 * File > Options "Maximum number of undos" trims the same stack.
 */
import { DEFAULT_EDITOR_HISTORY_DEPTH } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';

import type { CanvasSize } from '../types';
import type { EditorHistoryResult } from './useEditorHistory';
import { useEditorHistory } from './useEditorHistory';

function slide(color: string): PptxSlide {
	return {
		id: 'slide-1',
		rId: 'rId1',
		slideNumber: 1,
		backgroundColor: color,
		elements: [],
	} as unknown as PptxSlide;
}

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

let history: React.RefObject<EditorHistoryResult | null> | null = null;

function Harness({ maxHistoryEntries }: { maxHistoryEntries?: number }): React.ReactElement {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [slide('#000000')]);
	const [canvasSize, setCanvasSize] = useState<CanvasSize>({ width: 960, height: 540 });
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [templates, setTemplates] = useState({});
	const result = useEditorHistory({
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
		...(maxHistoryEntries === undefined ? {} : { maxHistoryEntries }),
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId: setTemplates,
		setSelectedElementId: noop<string | null>(),
		setSelectedElementIds: noop<string[]>(),
		setEditTemplateMode: noop<boolean>(),
		setHeaderFooter: noop<never>(),
	});
	const ref = useRef<EditorHistoryResult | null>(null);
	ref.current = result;
	history = ref;
	return <div />;
}

let root: Root | null = null;

function mount(props: { maxHistoryEntries?: number } = {}): () => EditorHistoryResult {
	const container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	act(() => {
		root?.render(<Harness {...props} />);
	});
	return () => history!.current!;
}

/** Commit `count` explicit, labelled slide edits, one render each. */
function commit(get: () => EditorHistoryResult, count: number): void {
	for (let i = 0; i < count; i++) {
		act(() => {
			get().commitSlides([slide(`#${String(i).padStart(6, '0')}`)], `edit ${i}`);
		});
	}
}

function undoAll(get: () => EditorHistoryResult): number {
	let undone = 0;
	while (get().canUndo) {
		act(() => get().handleUndo());
		undone += 1;
	}
	return undone;
}

afterEach(() => {
	act(() => root?.unmount());
	root = null;
	history = null;
});

describe('useEditorHistory on the shared history controller', () => {
	it('keeps the one shared default depth of undo steps', () => {
		const get = mount();
		commit(get, DEFAULT_EDITOR_HISTORY_DEPTH + 5);
		expect(get().undoLabel).toBe(`edit ${DEFAULT_EDITOR_HISTORY_DEPTH + 4}`);
		expect(undoAll(get)).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
		expect(get().canRedo).toBe(true);
	});

	it('applies the File > Options maximum to the same stack', () => {
		const get = mount({ maxHistoryEntries: 3 });
		commit(get, 6);
		expect(undoAll(get)).toBe(3);
	});
});
