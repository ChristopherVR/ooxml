// @vitest-environment happy-dom
/**
 * A live colour drag in the React inspector stays on the element it started
 * on, puts back exactly what it changed, and never overwrites another edit
 * (shared `createLiveColorInput`):
 *
 * - the selection moving mid-drag puts the first shape back and the pick does
 *   not land on the second, even when both show the same colour;
 * - an edit to another field of the dragged shape survives the commit;
 * - a drag cancelled back to the start leaves an element that had no colour of
 *   its own (no line, an inherited text colour) exactly as it was.
 *
 * The real `ShapeTextPanels` over the real `useEditorHistory` and
 * `useElementOperations`, with the live edit context `ViewerInspector` provides.
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

type Styled = PptxElement & { shapeStyle: ShapeStyle; textStyle?: TextStyle };

const A_STYLE: ShapeStyle = {
	fillMode: 'solid',
	fillColor: '#4472c4',
	fillColorRef: { scheme: 'accent1' },
	strokeColor: '#ed7d31',
	strokeWidth: 2,
	glowColor: '#ff0000',
	glowRadius: 8,
	glowOpacity: 0.5,
};

function shape(id: string, shapeStyle: ShapeStyle, textStyle?: TextStyle): PptxElement {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle,
		...(textStyle ? { textStyle } : {}),
	} as unknown as PptxElement;
}

interface HarnessApi {
	element(id: string): Styled;
	select(id: string): void;
	edit(id: string, patch: Partial<ShapeStyle>): void;
	canUndo(): boolean;
	recent(): string[];
}

let api: HarnessApi | null = null;

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
		element: (id) => latest.current.slides[0]!.elements.find((el) => el.id === id) as Styled,
		select: (id) => {
			setSelectedElementId(id);
			setSelectedElementIds([id]);
		},
		// Another hand (a collaborator) edits the shape: no undo step of this user's.
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
		canUndo: () => latest.current.history.canUndo,
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
		}),
		[history, ops, selectedElement],
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

function render(elements: PptxElement[]): void {
	act(() => root.render(<Harness elements={elements} />));
}

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

/** The colour input under the label whose caption reads `caption`. */
function inputUnder(caption: string): HTMLInputElement {
	const label = [...container.querySelectorAll('label')].find(
		(candidate) => candidate.querySelector('span')?.textContent?.trim() === caption,
	);
	const input = label?.querySelector<HTMLInputElement>('input[type="color"]');
	if (!input) {
		throw new Error(`no colour input under ${caption}`);
	}
	return input;
}

const fill = () => inputUnder('pptx.fill.fill');
const line = () => inputUnder('pptx.inspector.stroke');
const text = () =>
	container.querySelector<HTMLInputElement>('[data-pptx-text-card] input[type="color"]')!;

function fire(input: HTMLInputElement, type: 'input' | 'change' | 'blur', hex: string): void {
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

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', () => {
		const b = {
			...A_STYLE,
			fillColor: '#00aa55',
			fillColorRef: { scheme: 'accent6' },
		} as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		fire(fill(), 'input', '#330000');
		act(() => api!.select('b'));
		fire(fill(), 'change', '#330000');

		expect(api!.element('a').shapeStyle).toStrictEqual(A_STYLE);
		expect(api!.element('b').shapeStyle).toStrictEqual(b);
		expect(api!.canUndo()).toBeFalsy();
		expect(api!.recent()).toStrictEqual([]);
	});

	it('does not give the second shape the first one’s theme colour when both show the same hex', () => {
		const b = { ...A_STYLE, fillColorRef: { scheme: 'accent5' } } as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		fire(fill(), 'input', '#330000');
		fire(fill(), 'input', '#4472c4');
		act(() => api!.select('b'));
		fire(fill(), 'blur', '#4472c4');

		expect(api!.element('b').shapeStyle).toStrictEqual(b);
		expect(api!.element('a').shapeStyle).toStrictEqual(A_STYLE);
	});

	it('does not write the first shape’s effects onto the second', () => {
		const b = { ...A_STYLE, glowRadius: 3, glowOpacity: 0.9 } as ShapeStyle;
		render([shape('a', A_STYLE), shape('b', b)]);
		const glow = () => container.querySelector<HTMLInputElement>('[data-testid="fx-glow-color"]')!;
		fire(glow(), 'input', '#00ff00');
		fire(glow(), 'input', '#ff0000');
		act(() => api!.select('b'));
		fire(glow(), 'blur', '#ff0000');

		expect(api!.element('b').shapeStyle).toStrictEqual(b);
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit', () => {
		render([shape('a', A_STYLE)]);
		fire(fill(), 'input', '#330000');
		act(() => api!.edit('a', { strokeWidth: 5 }));
		fire(fill(), 'change', '#330000');

		expect(api!.element('a').shapeStyle).toMatchObject({ fillColor: '#330000', strokeWidth: 5 });
	});
});

describe('a cancelled drag on an element with no colour of its own', () => {
	it('leaves a shape without a line without one', () => {
		const style: ShapeStyle = { fillMode: 'solid', fillColor: '#4472c4' };
		render([shape('a', style)]);
		const start = line().value;
		fire(line(), 'input', '#330000');
		fire(line(), 'input', start);
		fire(line(), 'blur', start);

		expect(api!.element('a').shapeStyle).toStrictEqual(style);
		expect(api!.canUndo()).toBeFalsy();
	});

	it('leaves inherited text without a colour of its own', () => {
		render([shape('a', A_STYLE, {})]);
		const start = text().value;
		fire(text(), 'input', '#330000');
		fire(text(), 'change', start);

		expect(api!.element('a').textStyle).toStrictEqual({});
		expect(api!.canUndo()).toBeFalsy();
	});
});
