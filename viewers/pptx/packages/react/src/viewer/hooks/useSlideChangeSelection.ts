/**
 * A selection made in the same step as a slide change, and the effect that clears the selection
 * after every other slide change.
 *
 * The viewer drops the selection after a slide change, in an effect that runs after the render in
 * which the slide changed. Find goes to a match's slide and selects the match in that same step,
 * so the effect used to wipe the match's selection straight away. {@link
 * SlideChangeSelection.showSlideSelection} notes the slide the selection was made for, and
 * {@link useClearSelectionOnSlideChange} keeps it once, when the slide change it sees lands there.
 * Every other slide change (the slides pane, Page Down, the host's `setActiveSlideIndex`) clears
 * the selection as before.
 */
import { useCallback, useEffect, useRef } from 'react';
import type { Dispatch, MutableRefObject, SetStateAction } from 'react';

export interface SlideChangeSelection {
	/** The slide the last {@link showSlideSelection} selected on, until the next slide change. */
	slideSelectionRef: MutableRefObject<number | null>;
	/** Go to `slideIndex` and select `elementId` there (nothing, for null), across the slide change. */
	showSlideSelection: (slideIndex: number, elementId: string | null) => void;
}

interface SlideChangeSelectionSetters {
	setActiveSlideIndex: Dispatch<SetStateAction<number>>;
	setSelectedElementId: Dispatch<SetStateAction<string | null>>;
	setSelectedElementIds: Dispatch<SetStateAction<string[]>>;
}

export function useSlideChangeSelection({
	setActiveSlideIndex,
	setSelectedElementId,
	setSelectedElementIds,
}: SlideChangeSelectionSetters): SlideChangeSelection {
	const slideSelectionRef = useRef<number | null>(null);
	const showSlideSelection = useCallback(
		(slideIndex: number, elementId: string | null) => {
			slideSelectionRef.current = slideIndex;
			setActiveSlideIndex(slideIndex);
			setSelectedElementId(elementId);
			// One element, as a plain click selects it: the ids array is for multi-selections.
			setSelectedElementIds([]);
		},
		[setActiveSlideIndex, setSelectedElementId, setSelectedElementIds],
	);
	return { slideSelectionRef, showSlideSelection };
}

interface SlideChangeClearState extends SlideChangeSelectionSetters {
	slideSelectionRef: MutableRefObject<number | null>;
	setInlineEditingElementId: (id: string | null) => void;
}

/**
 * Clear the selection and any inline edit after a slide change, unless the selection was made for
 * the slide it landed on. Used by `PowerPointViewer` and by the headless building blocks.
 */
export function useClearSelectionOnSlideChange(
	state: Omit<SlideChangeClearState, 'setActiveSlideIndex'>,
	activeSlideIndex: number,
): void {
	useEffect(() => {
		const kept = state.slideSelectionRef.current === activeSlideIndex;
		state.slideSelectionRef.current = null;
		state.setInlineEditingElementId(null);
		if (kept) {
			return;
		}
		state.setSelectedElementId(null);
		state.setSelectedElementIds([]);
		// `activeSlideIndex` is the trigger; the setters and the ref never change.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [activeSlideIndex]);
}
