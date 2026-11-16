import { buildOutline, createOutlineViewSession } from 'ooxml-ui/pptx';
import type { OutlineEdit, OutlineRow } from 'ooxml-ui/pptx';
/**
 * Outline-view state and keyboard, kept out of the markup.
 *
 * The outline's rules and its edit session (`createOutlineViewSession`: refuse on a read-only
 * deck, apply, commit, hand the caret on) live in `ooxml-ui/pptx`; this hook derives the rows
 * from the deck, commits through the viewer's own `setSlides` and restores the caret once the
 * new rows have rendered.
 *
 * Undo works because of the `setSlides` + `bumpHistory` pair. React's history
 * hook watches slide state and skips its deep comparison when a cheap
 * structural hash is unchanged, and a text edit changes no slide or element
 * COUNT, so without the nonce bump an outline edit would be invisible to undo.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

import type { CanvasSize } from '../../types';

export interface UseOutlineViewInput {
	slides: PptxSlide[];
	canvasSize: CanvasSize;
	canEdit: boolean;
	setSlides: (slides: PptxSlide[]) => void;
	setActiveSlideIndex: (index: number) => void;
	/** Forces the history hook past its cheap-hash gate. See the module note. */
	bumpHistory: () => void;
}

export interface UseOutlineViewResult {
	rows: OutlineRow[];
	containerRef: React.RefObject<HTMLDivElement | null>;
	run: (edit: OutlineEdit) => void;
	onRowKeyDown: (event: React.KeyboardEvent, rowKey: string) => void;
}

export function useOutlineView(input: UseOutlineViewInput): UseOutlineViewResult {
	const latest = useRef(input);
	latest.current = input;
	const [session] = useState(() =>
		createOutlineViewSession({
			slides: () => latest.current.slides,
			canvasSize: () => latest.current.canvasSize,
			canEdit: () => latest.current.canEdit,
			commit: (slides, activeSlideIndex) => {
				latest.current.setSlides(slides);
				latest.current.setActiveSlideIndex(activeSlideIndex);
				latest.current.bumpHistory();
			},
		}),
	);
	const rows = useMemo(() => buildOutline(input.slides), [input.slides]);
	const containerRef = useRef<HTMLDivElement>(null);

	// The row that should take the caret may only exist once React has re-rendered with the new
	// deck, so the hand-off waits for the rows that render produced.
	useEffect(() => {
		session.restoreFocus(containerRef.current);
		// `rows` is not read in the body; it's the re-run trigger, so the hand-off happens only
		// once the rows an edit produced are in the DOM.
		// oxlint-disable-next-line react/exhaustive-effect-dependencies -- see comment above
	}, [rows, session]);

	return useMemo(
		() => ({
			rows,
			containerRef,
			run: (edit: OutlineEdit) => {
				session.run(edit);
			},
			onRowKeyDown: (event: React.KeyboardEvent, rowKey: string) =>
				session.handleKey(event, rowKey),
		}),
		[rows, session],
	);
}
