import {
	findReplaceSessionActions,
	INITIAL_FIND_REPLACE_SESSION,
	isEditorTextInputTarget,
	mapCustomizedEditorKey,
} from 'ooxml-ui/pptx';
import type { FindReplaceSession, FindResult, ResolvedKeyboardCustomization } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

// The match descriptor, the search / replace implementations and the session around them
// (`find-replace-session`) are shared with the other bindings; this hook keeps the session in
// React state.
export type { FindResult } from 'ooxml-ui/pptx';
interface UseFindReplaceInput {
	slides: PptxSlide[];
	mode: string;
	onSetActiveSlideIndex: (index: number) => void;
	onSetSelectedElementId: (id: string | null) => void;
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
	onSetActiveSlideIndex,
	onSetSelectedElementId,
	onUpdateSlides,
	onMarkDirty,
	keyboard,
}: UseFindReplaceInput): UseFindReplaceResult {
	const [findReplaceOpen, setFindReplaceOpen] = useState(false);
	const [session, setSession] = useState<FindReplaceSession>(INITIAL_FIND_REPLACE_SESSION);
	// The actions run long after the render that made them, so they read the latest slides and
	// callbacks, and the store reads a ref rather than this render's `session`.
	const latest = useRef({
		slides,
		onSetActiveSlideIndex,
		onSetSelectedElementId,
		onUpdateSlides,
		onMarkDirty,
	});
	latest.current = {
		slides,
		onSetActiveSlideIndex,
		onSetSelectedElementId,
		onUpdateSlides,
		onMarkDirty,
	};
	const latestSession = useRef(session);
	// One set of actions for the hook's lifetime (they read everything through the refs).
	const [actions] = useState(() =>
		findReplaceSessionActions(
			{
				get: () => latestSession.current,
				set: (next) => {
					latestSession.current = next;
					setSession(next);
				},
			},
			{
				slides: () => latest.current.slides,
				commit: (next) => {
					latest.current.onUpdateSlides(() => next as PptxSlide[]);
					latest.current.onMarkDirty();
				},
				navigate: ({ slideIndex, elementId }) => {
					latest.current.onSetActiveSlideIndex(slideIndex);
					latest.current.onSetSelectedElementId(elementId);
				},
			},
		),
	);

	// Closing the bar drops the matches. They also drive the canvas highlight
	// overlay, which is not gated on `findReplaceOpen`, so leaving them set
	// kept the match boxes painted on the slide after the bar was gone.
	useEffect(() => {
		if (!findReplaceOpen) {
			actions.close();
		}
	}, [findReplaceOpen, actions]);

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
	// The chord itself is decided by the shared editor keymap, not here. This
	// used to be a hand-rolled `event.key === 'f'` test, which is why Angular,
	// Svelte and Vanilla never got the shortcut at all: there was nothing in
	// shared for them to adopt. Note `key === 'f'` also missed Ctrl+Shift+F,
	// which `mapEditorKey` folds in via `toLowerCase()`. Ctrl+H (PowerPoint's
	// Find & Replace) opens the same single bar: it already renders both the
	// find and replace rows together, so there is no separate mode to enter.
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
			if (action !== 'find' && action !== 'findReplace') {
				return;
			}
			event.preventDefault();
			setFindReplaceOpen((prev) => !prev);
		};
		window.addEventListener('keydown', handler);
		return () => window.removeEventListener('keydown', handler);
	}, [mode, keyboard]);

	return {
		findReplaceOpen,
		setFindReplaceOpen,
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
