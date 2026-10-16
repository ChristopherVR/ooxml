// @vitest-environment happy-dom
/**
 * A Design gallery hover preview swaps the slides in for a moment. With the REAL
 * {@link useEditorHistory} the swap must be neither an undo step nor a dirty report while tracking
 * is held, the original coming back must read as no change, and a committed pick must still be
 * exactly one undoable step.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, useCallback, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';

import type { CanvasSize } from '../types';
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

interface HarnessApi {
	canUndo: () => boolean;
	background: () => string | undefined;
	dirtyReports: () => number;
	setBackground: (color: string) => void;
	hold: (held: boolean) => void;
	bumpAndCommit: (color: string) => void;
	undo: () => void;
}

let api: HarnessApi | null = null;
let dirty = 0;

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

function Harness(): React.ReactElement {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [slide('#FFFFFF')]);
	const [nonce, setNonce] = useState(0);
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
		pointerCommitNonce: nonce,
		onDirty: useCallback(() => {
			dirty += 1;
		}, []),
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
		canUndo: () => historyRef.current.canUndo,
		background: () => slidesRef.current[0]?.backgroundColor,
		dirtyReports: () => dirty,
		setBackground: (color) => setSlides([slide(color)]),
		hold: (held) => historyRef.current.holdTracking(held),
		bumpAndCommit: (color) => {
			setNonce((n) => n + 1);
			setSlides([slide(color)]);
			historyRef.current.markDirty();
		},
		undo: () => historyRef.current.handleUndo(),
	};
	return <div />;
}

let root: Root | null = null;

function mount(): HarnessApi {
	dirty = 0;
	const container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	act(() => {
		root?.render(<Harness />);
	});
	return api!;
}

afterEach(() => {
	act(() => root?.unmount());
	root = null;
	api = null;
});

describe('holdTracking', () => {
	it('records nothing while a preview shows or when the original comes back', () => {
		const harness = mount();
		const original = harness.background();
		act(() => {
			harness.hold(true);
			harness.setBackground('#000000');
		});
		expect(harness.background()).toBe('#000000');
		expect(harness.canUndo()).toBeFalsy();
		act(() => {
			harness.setBackground('#FF0000');
		});
		act(() => {
			harness.setBackground(original!);
			harness.hold(false);
		});
		expect(harness.background()).toBe(original);
		expect(harness.canUndo()).toBeFalsy();
		expect(harness.dirtyReports()).toBe(0);
	});

	it('still records a committed pick as one undo step, and undo puts the original back', () => {
		const harness = mount();
		act(() => {
			harness.hold(true);
			harness.setBackground('#000000');
		});
		act(() => {
			harness.setBackground('#FFFFFF');
			harness.hold(false);
			harness.bumpAndCommit('#123456');
		});
		expect(harness.background()).toBe('#123456');
		expect(harness.canUndo()).toBeTruthy();
		expect(harness.dirtyReports()).toBeGreaterThan(0);
		act(() => harness.undo());
		expect(harness.background()).toBe('#FFFFFF');
		expect(harness.canUndo()).toBeFalsy();
	});
});
