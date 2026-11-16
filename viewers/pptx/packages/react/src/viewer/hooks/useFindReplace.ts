import {
	createFindReplaceController,
	isEditorTextInputTarget,
	mapCustomizedEditorKey,
} from 'ooxml-ui/pptx';
import type { FindResult, ResolvedKeyboardCustomization } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

import { useController } from './useController';

// The bar's open state, its mode and the search session are the shared
// `createFindReplaceController`; this hook reads it through `useController`.
export type { FindResult } from 'ooxml-ui/pptx';
interface UseFindReplaceInput {
	slides: PptxSlide[];
	mode: string;
	/**
	 * Go to a match's slide and select its element there (null for a speaker-notes match), in one
	 * step: the viewer's `showSlideSelection`, which keeps the selection across the slide change.
	 */
	onShowMatch: (slideIndex: number, elementId: string | null) => void;
	onUpdateSlides: (updater: (slides: PptxSlide[]) => PptxSlide[]) => void;
	onMarkDirty: () => void;
	/** Host keyboard customisation (disabled / remapped Find shortcuts). */
	keyboard?: ResolvedKeyboardCustomization;
}

interface UseFindReplaceResult {
	findReplaceOpen: boolean;
	setFindReplaceOpen: (open: boolean) => void;
	findQuery: string;
	setFindQuery: (query: string) => void;
	replaceQuery: string;
	setReplaceQuery: (query: string) => void;
	findMatchCase: boolean;
	findIncludeNotes: boolean;
	findIncludeChartTitles: boolean;
	findIncludeChartData: boolean;
	setFindIncludeChartTitles: (value: boolean) => void;
	setFindIncludeChartData: (value: boolean) => void;
	setFindIncludeNotes: (value: boolean) => void;
	setFindMatchCase: (matchCase: boolean) => void;
	findResults: readonly FindResult[];
	findResultIndex: number;
	performFind: () => void;
	navigateFindResult: (direction: 1 | -1) => void;
	handleReplace: () => void;
	handleReplaceAll: () => void;
}

export function useFindReplace({
	slides,
	mode,
	onShowMatch,
	onUpdateSlides,
	onMarkDirty,
	keyboard,
}: UseFindReplaceInput): UseFindReplaceResult {
	// The actions run long after the render that made them, so they read the latest slides and
	// callbacks through a ref.
	const latest = useRef({ slides, onShowMatch, onUpdateSlides, onMarkDirty });
	latest.current = { slides, onShowMatch, onUpdateSlides, onMarkDirty };
	// One controller for the hook's lifetime (it reads everything through the ref).
	const [actions] = useState(() =>
		createFindReplaceController({
			slides: () => latest.current.slides,
			commit: (next) => {
				latest.current.onUpdateSlides(() => next as PptxSlide[]);
				latest.current.onMarkDirty();
			},
			navigate: ({ slideIndex, elementId }) => latest.current.onShowMatch(slideIndex, elementId),
		}),
	);
	const { open: findReplaceOpen, session } = useController(actions);

	// React's slides land a render after a replace commits them, so the session
	// searches again here, once they have. It does nothing after an ordinary
	// edit, so editing with the bar open does not jump back to a match.
	useEffect(() => {
		actions.slidesChanged();
		// `slides` is the trigger (the actions read the latest slides themselves); `actions`
		// never changes.
		// oxlint-disable-next-line react/exhaustive-effect-dependencies -- see comment above
	}, [slides, actions]);

	const handlers = useMemo(
		() => ({
			setFindReplaceOpen: actions.setOpen,
			setFindQuery: (query: string) => actions.search({ query }),
			setReplaceQuery: actions.setReplacement,
			setFindMatchCase: (matchCase: boolean) => actions.search({ options: { matchCase } }),
			setFindIncludeNotes: (includeNotes: boolean) => actions.search({ options: { includeNotes } }),
			setFindIncludeChartTitles: (includeChartTitles: boolean) =>
				actions.search({ options: { includeChartTitles } }),
			setFindIncludeChartData: (includeChartData: boolean) =>
				actions.search({ options: { includeChartData } }),
			performFind: () => actions.search(),
			navigateFindResult: (direction: 1 | -1) =>
				direction > 0 ? actions.next() : actions.previous(),
			handleReplace: () => {
				actions.replaceCurrent();
			},
			handleReplaceAll: () => {
				actions.replaceAll();
			},
		}),
		[actions],
	);

	// ── Keyboard shortcut: Ctrl/Cmd+F or Ctrl/Cmd+H toggles the find bar ──
	// The chord is decided by the shared editor keymap and the toggle by the
	// shared controller (Ctrl+H switches an open Ctrl+F bar to its replace mode).
	useEffect(() => {
		const handler = (event: KeyboardEvent) => {
			const { action } = mapCustomizedEditorKey(
				event,
				{
					canEdit: mode === 'edit',
					isTextInputTarget: isEditorTextInputTarget(event.target),
				},
				keyboard,
			);
			if (action === 'find' || action === 'findReplace') {
				event.preventDefault();
				actions.handleShortcut(action);
			}
		};
		window.addEventListener('keydown', handler);
		return () => window.removeEventListener('keydown', handler);
	}, [mode, keyboard, actions]);

	return {
		findReplaceOpen,
		findQuery: session.query,
		replaceQuery: session.replacement,
		findMatchCase: session.options.matchCase,
		findIncludeNotes: session.options.includeNotes,
		findIncludeChartTitles: session.options.includeChartTitles,
		findIncludeChartData: session.options.includeChartData,
		findResults: session.matches,
		findResultIndex: session.current,
		...handlers,
	};
}
