/**
 * colour-drag.test-support.tsx: the harness the inspector's live colour drag
 * suites share. The real `ShapeTextPanels` over the real `useEditorHistory` and
 * `useElementOperations`, with the live edit context `ViewerInspector`
 * provides, mounted fresh for each test with element `a` selected.
 *
 * Not a `.test.tsx` file on purpose: a shared harness is not a suite. Each
 * suite keeps its own `vi.mock` calls, which apply to what this module imports.
 */
import type { PptxElement, PptxSlide, ShapeStyle, TextStyle } from 'pptx-viewer-core';
import React, { act, useCallback, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, vi } from 'vitest';

import { useEditorHistory } from '../../hooks/useEditorHistory';
import { useElementOperations } from '../../hooks/useElementOperations';
import type { CanvasSize, ElementContextMenuState } from '../../types';
import { LiveEditProvider } from './LiveEditContext';
import { RecentColorsProvider } from './RecentColorsContext';
import { ShapeTextPanels } from './ShapeTextPanels';

export type Styled = PptxElement & { shapeStyle: ShapeStyle; textStyle?: TextStyle };

export interface HarnessApi {
	/** Element `id` (`a`, the one selected first, when omitted) as the deck holds it now. */
	element(id?: string): Styled;
	select(id: string): void;
	/** Another hand (a collaborator) edits the shape: no undo step of this user's. */
	edit(id: string, patch: Partial<ShapeStyle>): void;
	editText(id: string, patch: Partial<TextStyle>): void;
	canUndo(): boolean;
	undo(): void;
	recent(): string[];
}

/** The mounted harness (a live binding: set while a test's harness is rendered). */
export let api: HarnessApi | null = null;
let container: HTMLDivElement;
let root: Root;

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

function Harness({ elements }: { elements: PptxElement[] }): React.ReactElement {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [
		{ id: 's1', rId: 'rId1', slideNumber: 1, elements } as unknown as PptxSlide,
	]);
	const [canvasSize, setCanvasSize] = useState<CanvasSize>({ width: 960, height: 540 });
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [templateElementsBySlideId, setTemplateElementsBySlideId] = useState<
		Record<string, PptxElement[]>
	>({});
	const [selectedElementId, setSelectedElementId] = useState<string | null>('a');
	const [selectedElementIds, setSelectedElementIds] = useState<string[]>(['a']);
	const [recent, setRecent] = useState<string[]>([]);

	const history = useEditorHistory({
		slides,
		canvasSize,
		activeSlideIndex,
		templateElementsBySlideId,
		selectedElementId,
		selectedElementIds,
		editTemplateMode: false,
		headerFooter: {} as never,
		loading: false,
		error: null,
		hasActivePointerInteraction: useCallback(() => false, []),
		pointerCommitNonce: 0,
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId,
		setSelectedElementId,
		setSelectedElementIds,
		setEditTemplateMode: noop<boolean>(),
		setHeaderFooter: () => {},
	});
	const activeSlide = slides[activeSlideIndex];
	const selectedElement = useMemo(
		() => activeSlide?.elements.find((el) => el.id === selectedElementId) ?? null,
		[activeSlide, selectedElementId],
	);
	const ops = useElementOperations({
		slides,
		activeSlide,
		activeSlideIndex,
		selectedElement,
		selectedElementId,
		editTemplateMode: false,
		templateElements: [],
		history,
		setSlides,
		setTemplateElementsBySlideId,
		setSelectedElementId,
		setSelectedElementIds,
		setInlineEditingElementId: noop<string | null>(),
		setContextMenuState: noop<ElementContextMenuState | null>(),
		inlineEditingElementId: null,
		inlineEditingText: '',
	});

	const latest = useRef({ slides, history, recent });
	latest.current = { slides, history, recent };
	api = {
		element: (id = 'a') => latest.current.slides[0]!.elements.find((el) => el.id === id) as Styled,
		select: (id) => {
			setSelectedElementId(id);
			setSelectedElementIds([id]);
		},
		edit: (id, patch) =>
			setSlides((prev) =>
				prev.map((slide) => ({
					...slide,
					elements: slide.elements.map((el) =>
						el.id === id
							? ({ ...el, shapeStyle: { ...(el as Styled).shapeStyle, ...patch } } as PptxElement)
							: el,
					),
				})),
			),
		editText: (id, patch) =>
			setSlides((prev) =>
				prev.map((slide) => ({
					...slide,
					elements: slide.elements.map((el) =>
						el.id === id
							? ({ ...el, textStyle: { ...(el as Styled).textStyle, ...patch } } as PptxElement)
							: el,
					),
				})),
			),
		canUndo: () => latest.current.history.canUndo,
		undo: () => latest.current.history.handleUndo(),
		recent: () => latest.current.recent,
	};
	const pushColor = useCallback(
		(hex: string) => setRecent((list) => [hex, ...list.filter((c) => c !== hex)]),
		[],
	);
	const recentValue = useMemo(() => ({ recentColors: recent, pushColor }), [recent, pushColor]);
	const liveEdit = useMemo(
		() => ({
			element: selectedElement,
			previewElement: (id: string, patch: Partial<PptxElement>) =>
				history.withoutUndoStep(() => ops.updateElementById(id, patch)),
			elementById: (id: string) => activeSlide?.elements.find((el) => el.id === id),
		}),
		[history, ops, selectedElement, activeSlide],
	);

	return (
		<LiveEditProvider value={liveEdit}>
			<RecentColorsProvider value={recentValue}>
				{selectedElement && (
					<ShapeTextPanels
						selectedElement={selectedElement}
						canEdit
						onUpdateElement={ops.updateSelectedElement}
						onUpdateElementStyle={ops.updateSelectedShapeStyle}
						onUpdateTextStyle={ops.updateSelectedTextStyle}
					/>
				)}
			</RecentColorsProvider>
		</LiveEditProvider>
	);
}

/** A fresh container and root for each test of the calling suite, unmounted after it. */
export function setUpColourDragHarness(): void {
	beforeEach(() => {
		vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
		container = document.createElement('div');
		document.body.appendChild(container);
		root = createRoot(container);
	});

	afterEach(() => {
		act(() => root.unmount());
		container.remove();
		api = null;
		vi.unstubAllGlobals();
	});
}

/** Mount the panels over a slide of `elements`. */
export function renderHarness(elements: PptxElement[]): void {
	act(() => root.render(<Harness elements={elements} />));
}

export type Control = 'fill' | 'line' | 'text' | 'glow';

/** The colour input under the label whose caption reads `caption`. */
function inputUnder(caption: string): HTMLInputElement | null {
	const label = [...container.querySelectorAll('label')].find(
		(candidate) => candidate.querySelector('span')?.textContent?.trim() === caption,
	);
	return label?.querySelector<HTMLInputElement>('input[type="color"]') ?? null;
}

const INPUTS: Record<Control, () => HTMLInputElement | null> = {
	fill: () => inputUnder('pptx.fill.fill'),
	line: () => inputUnder('pptx.inspector.stroke'),
	text: () =>
		container.querySelector<HTMLInputElement>('[data-pptx-text-card] input[type="color"]'),
	glow: () => container.querySelector<HTMLInputElement>('[data-testid="fx-glow-color"]'),
};

/** The `control` colour input as rendered now. */
export function colorInput(control: Control): HTMLInputElement {
	const input = INPUTS[control]();
	if (!input) {
		throw new Error(`no ${control} colour input`);
	}
	return input;
}

/** Fire a native colour-input event the way the browser does, so React's value tracker sees it. */
export function fire(
	input: HTMLInputElement,
	type: 'input' | 'change' | 'blur',
	hex: string,
): void {
	act(() => {
		Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, hex);
		input.dispatchEvent(
			type === 'blur'
				? new FocusEvent('focusout', { bubbles: true })
				: new Event(type, { bubbles: true }),
		);
		if (type === 'blur') {
			input.dispatchEvent(new FocusEvent('blur'));
		}
	});
}
