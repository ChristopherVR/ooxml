// @vitest-environment happy-dom
/**
 * Apply all, through real React state: every proposal applies in one tick, before any re-render,
 * so each must see what the previous one wrote. All of them must reach the deck, each as its own
 * undo step.
 */
import { checkPptxAiBridgeReadsItsWrites, ProposalStore } from 'ooxml-ui/pptx/ai';
import type { PptxAiBridge, ProposalCommit } from 'ooxml-ui/pptx/ai';
import type {
	ParsedTableStyleMap,
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxCustomShow,
	PptxData,
	PptxElement,
	PptxHandler,
	PptxPresentationProperties,
	PptxSection,
	PptxSlide,
	PptxTagCollection,
	PptxTheme,
	PptxThemeColorScheme,
	PptxViewProperties,
} from 'pptx-viewer-core';
import React, { act, useCallback, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import type { EditorHistoryResult } from '../useEditorHistory';
import { useEditorHistory } from '../useEditorHistory';
import { useAiBridge } from './useAiBridge';

function element(id: string, text: string): PptxElement {
	return { id, type: 'text', x: 0, y: 0, width: 10, height: 10, text } as unknown as PptxElement;
}

function initialSlides(): PptxSlide[] {
	return [
		{ id: 's0', slideNumber: 1, elements: [element('a', 'A'), element('b', 'B')] },
	] as unknown as PptxSlide[];
}

const INITIAL_THEME = {
	colorScheme: { accent1: '#000001', accent2: '#000002' },
	fontScheme: { majorFont: { latin: 'Serif' } },
} as unknown as PptxTheme;

interface Rendered {
	bridge: PptxAiBridge;
	history: EditorHistoryResult;
	slides: PptxSlide[];
	theme: PptxTheme | undefined;
	core: PptxCoreProperties | undefined;
	replaceSlides: (slides: PptxSlide[]) => void;
	/** What a load does: a new deck load and handler, then the new deck's state. */
	openDeck: (slides: PptxSlide[], core: PptxCoreProperties | undefined) => void;
	/** What a Slide Master view command or a slide paste does: a new handler on the same load. */
	swapHandler: () => void;
}

let rendered: Rendered;
const handlerRef: { current: PptxHandler | null } = { current: null };

const noop = (): void => {};

function Harness(): null {
	const [slides, setSlides] = useState(initialSlides);
	const [theme, setTheme] = useState<PptxTheme | undefined>(INITIAL_THEME);
	const [canvasSize, setCanvasSize] = useState({ width: 960, height: 540 });
	const [sections, setSections] = useState<PptxSection[]>([]);
	const [customShows, setCustomShows] = useState<PptxCustomShow[]>([]);
	const [presentationProperties, setPresentationProperties] = useState<PptxPresentationProperties>(
		{},
	);
	const [customProperties, setCustomProperties] = useState<PptxCustomProperty[]>([]);
	const [coreProperties, setCoreProperties] = useState<PptxCoreProperties | undefined>({});
	const [appProperties, setAppProperties] = useState<PptxAppProperties | undefined>();
	const [viewProperties, setViewProperties] = useState<PptxViewProperties | undefined>();
	const [tableStyleMap, setTableStyleMap] = useState<ParsedTableStyleMap | undefined>();
	const [tableStylesDefaultId, setTableStylesDefaultId] = useState<string | undefined>();
	const [tagCollections, setTagCollections] = useState<PptxTagCollection[]>([]);
	const [templateElementsBySlideId, setTemplateElementsBySlideId] = useState<
		Record<string, PptxElement[]>
	>({});
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [deckLoad, setDeckLoad] = useState(1);
	const history = useEditorHistory({
		slides,
		canvasSize,
		activeSlideIndex,
		templateElementsBySlideId,
		selectedElementId: null,
		selectedElementIds: [],
		editTemplateMode: false,
		headerFooter: {} as never,
		loading: false,
		error: null,
		hasActivePointerInteraction: useCallback(() => false, []),
		pointerCommitNonce: 0,
		theme,
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId,
		setSelectedElementId: noop,
		setSelectedElementIds: noop,
		setEditTemplateMode: noop,
		setHeaderFooter: noop,
	});
	const bridge = useAiBridge({
		slides,
		activeSlideIndex,
		canvasSize,
		theme,
		selectedElementId: null,
		selectedElementIds: [],
		pinnedFocus: null,
		handlerRef,
		deckLoad,
		sections,
		customShows,
		presentationProperties,
		customProperties,
		coreProperties,
		appProperties,
		viewProperties,
		tableStyleMap,
		tableStylesDefaultId,
		tagCollections,
		setCanvasSize,
		setSections,
		setCustomShows,
		setPresentationProperties,
		setCustomProperties,
		setCoreProperties,
		setAppProperties,
		setViewProperties,
		setTableStyleMap,
		setTableStylesDefaultId,
		setTagCollections,
		setSlides,
		setActiveSlideIndex,
		applySelection: noop,
		recordSlidesStep: history.recordSlidesCommit,
		markDirty: history.markDirty,
		applyThemeUpdates: (updates) => setTheme((previous) => ({ ...previous, ...updates })),
	});
	rendered = {
		bridge,
		history,
		slides,
		theme,
		core: coreProperties,
		replaceSlides: setSlides,
		openDeck(nextSlides, core) {
			handlerRef.current = {} as PptxHandler;
			setDeckLoad((load) => load + 1);
			setSlides(nextSlides);
			setCoreProperties(core);
		},
		swapHandler() {
			handlerRef.current = {} as PptxHandler;
		},
	};
	return null;
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	handlerRef.current = {} as PptxHandler;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	act(() => root.render(<Harness />));
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function propose(
	store: ProposalStore,
	label: string,
	commit: ProposalCommit,
	change: (deck: PptxData) => void,
): void {
	const before = structuredClone(rendered.bridge.getDeckData?.() as PptxData);
	const after = structuredClone(before);
	change(after);
	store.propose({ label, commit, before, after, requiresApproval: false });
}

function setText(deck: PptxData, index: number, text: string): void {
	const slide = deck.slides[0];
	if (slide) {
		slide.elements[index] = { ...slide.elements[index], text } as PptxElement;
	}
}

function setAccent(deck: PptxData, accent: 'accent1' | 'accent2', color: string): void {
	deck.theme = {
		...deck.theme,
		colorScheme: { ...deck.theme?.colorScheme, [accent]: color } as PptxThemeColorScheme,
	};
}

const textOf = (index: number): unknown =>
	(rendered.slides[0]?.elements[index] as { text?: string } | undefined)?.text;

describe('useAiBridge Apply all', () => {
	it('keeps every proposal applied in one tick', () => {
		const store = new ProposalStore(rendered.bridge);
		propose(store, 'Text A', 'slides', (deck) => setText(deck, 0, 'NEW A'));
		propose(store, 'Text B', 'slides', (deck) => setText(deck, 1, 'NEW B'));
		propose(store, 'Accent 1', 'theme', (deck) => setAccent(deck, 'accent1', '#a11111'));
		propose(store, 'Accent 2', 'theme', (deck) => setAccent(deck, 'accent2', '#a22222'));
		propose(store, 'Title', 'deck', (deck) => {
			deck.coreProperties = { ...deck.coreProperties, title: 'New title' };
		});
		propose(store, 'Author', 'deck', (deck) => {
			deck.coreProperties = { ...deck.coreProperties, creator: 'New author' };
		});

		let applied = 0;
		act(() => {
			applied = store.acceptAll();
		});

		expect(applied).toBe(6);
		expect(store.list()).toEqual([]);
		expect([textOf(0), textOf(1)]).toEqual(['NEW A', 'NEW B']);
		expect(rendered.theme?.colorScheme).toMatchObject({ accent1: '#a11111', accent2: '#a22222' });
		expect(rendered.core).toMatchObject({ title: 'New title', creator: 'New author' });
	});

	it('still reports a real conflict between proposals for the same text', () => {
		const store = new ProposalStore(rendered.bridge);
		propose(store, 'First', 'slides', (deck) => setText(deck, 0, 'FIRST'));
		propose(store, 'Second', 'slides', (deck) => setText(deck, 0, 'SECOND'));

		act(() => {
			store.acceptAll();
		});

		expect(textOf(0)).toBe('FIRST');
		expect(store.list().map((view) => [view.label, Boolean(view.conflict)])).toEqual([
			['Second', true],
		]);
	});

	it('makes each applied proposal its own undo step', () => {
		const store = new ProposalStore(rendered.bridge);
		propose(store, 'Text A', 'slides', (deck) => setText(deck, 0, 'NEW A'));
		propose(store, 'Text B', 'slides', (deck) => setText(deck, 1, 'NEW B'));
		act(() => {
			store.acceptAll();
		});
		expect(rendered.history.undoLabel).toBe('Text B');

		act(() => rendered.history.handleUndo());
		expect([textOf(0), textOf(1)]).toEqual(['NEW A', 'B']);
		expect(rendered.history.undoLabel).toBe('Text A');

		act(() => rendered.history.handleUndo());
		expect([textOf(0), textOf(1)]).toEqual(['A', 'B']);
	});

	it('meets the bridge contract before and after the writes render', () => {
		let failures: string[] = [];
		act(() => {
			failures = checkPptxAiBridgeReadsItsWrites(rendered.bridge);
		});
		expect(failures).toEqual([]);
		expect(rendered.slides[0]?.elements[0]).toMatchObject({ x: 1111, y: 2222 });
		expect(rendered.theme?.fontScheme?.majorFont?.latin).toBe('Bridge Contract Heading');
		expect(rendered.bridge.getTheme()?.colorScheme?.accent2).toBe('#A22222');
	});

	it('lets a render with newer state replace what was written', () => {
		act(() => {
			rendered.bridge.applySlidesUpdate((slides) => slides, 'Pending');
		});
		const newer = initialSlides();
		act(() => rendered.replaceSlides(newer));
		expect(rendered.bridge.getSlides()).toEqual(newer);
	});

	it('keeps what it wrote, and staged proposals, when an edit replaces the handler', () => {
		const store = new ProposalStore(rendered.bridge);
		propose(store, 'Author', 'deck', (deck) => {
			deck.coreProperties = { ...deck.coreProperties, creator: 'Staged author' };
		});
		act(() => {
			rendered.bridge.applyDeckData?.((deck) => {
				deck.coreProperties = { ...deck.coreProperties, title: 'Written title' };
				return deck;
			}, 'Title');
			rendered.swapHandler();
			expect(rendered.bridge.getDeckData?.()?.coreProperties?.title).toBe('Written title');
		});

		expect(store.list().map((view) => view.label)).toStrictEqual(['Author']);
		act(() => {
			store.acceptAll();
		});
		expect(rendered.core).toMatchObject({ title: 'Written title', creator: 'Staged author' });
	});

	it('forgets what it wrote once another deck opens', () => {
		const store = new ProposalStore(rendered.bridge);
		act(() => rendered.openDeck(initialSlides(), undefined));
		propose(store, 'Title', 'deck', (deck) => {
			deck.coreProperties = { ...deck.coreProperties, title: 'First deck title' };
		});
		act(() => {
			store.acceptAll();
		});
		expect(rendered.core).toMatchObject({ title: 'First deck title' });

		act(() => rendered.openDeck(initialSlides(), undefined));

		expect(rendered.core).toBeUndefined();
		expect(rendered.bridge.getDeckData?.()?.coreProperties?.title).toBeUndefined();
	});
});
