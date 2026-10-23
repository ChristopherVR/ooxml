// @vitest-environment happy-dom
/**
 * A drag through one of the inspector's native colour pickers (fill, line,
 * text, or an effects colour field) previews every shade and is ONE undo step
 * and ONE recent colour (shared `createLiveColorInput`), as in the other four
 * bindings.
 *
 * Every shade used to go through the inspector's normal edit path, which arms
 * the undo tracker (`markDirty`), so a drag through twenty shades took twenty
 * presses of Undo. Wired as the viewer wires it: the real `ShapeTextPanels`
 * over the real `useEditorHistory` and `useElementOperations`, with the live
 * edit context `ViewerInspector` provides.
 */
import type { PptxElement, PptxSlide, ShapeStyle, TextStyle } from 'pptx-viewer-core';
import React, { act, useCallback, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useEditorHistory } from '../../hooks/useEditorHistory';
import { useElementOperations } from '../../hooks/useElementOperations';
import type { CanvasSize, ElementContextMenuState } from '../../types';
import { LiveEditProvider } from './LiveEditContext';
import { RecentColorsProvider } from './RecentColorsContext';
import { ShapeTextPanels } from './ShapeTextPanels';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string, fallback?: unknown) => (typeof fallback === 'string' ? fallback : key),
	}),
}));

type Styled = PptxElement & { shapeStyle: ShapeStyle; textStyle: TextStyle };

function shape(): PptxElement {
	return {
		type: 'shape',
		id: 'a',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle: {
			fillMode: 'solid',
			fillColor: '#4472c4',
			fillColorRef: { scheme: 'accent1' },
			strokeColor: '#ed7d31',
			strokeColorRef: { scheme: 'accent2' },
			strokeWidth: 2,
			glowColor: '#ff0000',
			glowRadius: 8,
			glowOpacity: 0.5,
		},
		textStyle: { color: '#112233', colorRef: { scheme: 'tx1' } },
	} as unknown as PptxElement;
}

interface HarnessApi {
	element(): Styled;
	canUndo(): boolean;
	undo(): void;
	recent(): string[];
}

let api: HarnessApi | null = null;

function noop<T>(): React.Dispatch<React.SetStateAction<T>> {
	return () => {};
}

function Harness(): React.ReactElement {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [
		{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [shape()] } as unknown as PptxSlide,
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
		setHeaderFooter: noop<never>(),
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
		element: () => latest.current.slides[0]!.elements[0] as Styled,
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

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	act(() => root.render(<Harness />));
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	api = null;
	vi.unstubAllGlobals();
});

type Control = 'fill' | 'line' | 'text' | 'glow';

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

const COLOR: Record<Control, (el: Styled) => string | undefined> = {
	fill: (el) => el.shapeStyle.fillColor,
	line: (el) => el.shapeStyle.strokeColor,
	text: (el) => el.textStyle.color,
	glow: (el) => el.shapeStyle.glowColor,
};
const REF: Record<Control, (el: Styled) => unknown> = {
	fill: (el) => el.shapeStyle.fillColorRef,
	line: (el) => el.shapeStyle.strokeColorRef,
	text: (el) => el.textStyle.colorRef,
	glow: (el) => el.shapeStyle.glowRadius,
};

/** Fire a native colour-input event the way the browser does, so React's value tracker sees it. */
function fire(control: Control, type: 'input' | 'change' | 'blur', hex: string): void {
	const input = INPUTS[control]();
	if (!input) {
		throw new Error(`no ${control} colour input`);
	}
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

/** Set the value and fire `input` then `change` in one act, so React renders only after both. */
function pick(control: Control, hex: string): void {
	const input = INPUTS[control]();
	if (!input) {
		throw new Error(`no ${control} colour input`);
	}
	act(() => {
		Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, hex);
		input.dispatchEvent(new Event('input', { bubbles: true }));
		input.dispatchEvent(new Event('change', { bubbles: true }));
	});
}

const SHADES = ['#330000', '#660000', '#990000'];

describe.each(['fill', 'line', 'text', 'glow'] as const)('the %s colour picker', (control) => {
	const color = COLOR[control];
	const themeRef = REF[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', () => {
		fire(control, 'input', '#330000');

		expect(color(api!.element())).toBe('#330000');
		expect(api!.canUndo()).toBeFalsy();
		expect(api!.recent()).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', () => {
		const before = api!.element();
		for (const shade of SHADES) {
			fire(control, 'input', shade);
			expect(color(api!.element())).toBe(shade);
		}
		fire(control, 'change', '#990000');

		expect(color(api!.element())).toBe('#990000');
		if (control !== 'glow') {
			// A native pick is a custom colour: the theme ref goes.
			expect(themeRef(api!.element())).toBeUndefined();
		}
		expect(api!.recent()).toStrictEqual(['#990000']);

		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
		expect(api!.canUndo()).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', () => {
		const before = api!.element();
		fire(control, 'input', '#330000');
		fire(control, 'input', color(before)!);
		fire(control, 'change', color(before)!);

		expect(api!.canUndo()).toBeFalsy();
		// The `change` is still a pick: its colour counts as recent, though nothing was edited.
		expect(api!.recent()).toStrictEqual([color(before)]);
		expect(color(api!.element())).toBe(color(before));
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
	});

	it('commits a pick whose input and change arrive in one turn as one undo step', () => {
		const before = api!.element();
		pick(control, '#00ff00');

		expect(color(api!.element())).toBe('#00ff00');
		expect(api!.recent()).toStrictEqual(['#00ff00']);
		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(api!.canUndo()).toBeFalsy();
	});

	it('counts a pick of the colour already there as a recent colour, with no undo step', () => {
		const before = api!.element();
		pick(control, color(before)!);

		expect(api!.recent()).toStrictEqual([color(before)]);
		expect(api!.canUndo()).toBeFalsy();
		expect(themeRef(api!.element())).toStrictEqual(themeRef(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', () => {
		const before = api!.element();
		fire(control, 'input', '#330000');
		fire(control, 'blur', '#330000');

		expect(api!.recent()).toStrictEqual(['#330000']);
		act(() => api!.undo());
		expect(color(api!.element())).toBe(color(before));
		expect(api!.canUndo()).toBeFalsy();
	});
});
