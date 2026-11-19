/**
 * useVirtualizedSlides: Virtual scrolling for the slide panel sidebar.
 *
 * Calculates the visible range of slides based on the scroll container's
 * scroll position and viewport height, then returns only the indices
 * that should be rendered. An overscan buffer ensures smooth scrolling
 * by pre-rendering items just outside the viewport.
 *
 * @module useVirtualizedSlides
 */
import { createDeferredRows, createVirtualWindow, DEFAULT_VIRTUAL_OVERSCAN } from 'ooxml-ui/pptx';
import type { DeferredRows, VirtualizedRange, VirtualWindow } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';

export { computeVirtualRange } from 'ooxml-ui/pptx';
export type { VirtualizedRange } from 'ooxml-ui/pptx';

const useIsomorphicLayoutEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect;

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

export interface VirtualizedSlidesOptions {
	/** Total number of slide items. */
	totalItems: number;
	/** Estimated height of each slide item in pixels. */
	itemHeight: number;
	/** Number of extra items to render above/below the viewport. */
	overscan?: number;
}

export interface VirtualizedSlidesResult extends VirtualizedRange {
	/** Callback ref for the scroll container; re-wires scroll tracking whenever it mounts. */
	scrollContainerRef: (el: HTMLDivElement | null) => void;
	/** Call this to scroll a specific index into view. */
	scrollToIndex: (index: number, behavior?: ScrollBehavior) => void;
	/**
	 * Report the keys of the rows now rendered. A row new to the window is
	 * listed in `readyRows` once scrolling pauses, so its thumbnail is built
	 * then rather than during the scroll.
	 */
	syncRows: (keys: readonly string[]) => void;
	/** Rendered rows whose deferred thumbnail may be built. */
	readyRows: ReadonlySet<string>;
}

/* ------------------------------------------------------------------ */
/*  Hook                                                              */
/* ------------------------------------------------------------------ */

export function useVirtualizedSlides({
	totalItems,
	itemHeight,
	overscan = DEFAULT_VIRTUAL_OVERSCAN,
}: VirtualizedSlidesOptions): VirtualizedSlidesResult {
	// The shared window reads scroll and resize once per frame and reports only
	// a change of rendered rows, so scrolling re-renders the rail per entering
	// row rather than per scroll event.
	const [range, setRange] = useState<VirtualizedRange | null>(null);
	const viewRef = useRef<VirtualWindow | null>(null);
	viewRef.current ??= createVirtualWindow({ totalItems, itemHeight, overscan }, setRange);
	const view = viewRef.current;
	const [readyRows, setReadyRows] = useState<ReadonlySet<string>>(() => new Set());
	const rowsRef = useRef<DeferredRows | null>(null);
	rowsRef.current ??= createDeferredRows(view, setReadyRows);
	const deferredRows = rowsRef.current;
	// The container is tracked in state, not only in a ref: the rail renders it
	// only while open, so it can mount after this hook's first effect. A
	// mount-once effect then never saw it, the viewport stayed 0 px tall and
	// the rail never rendered past its first dozen thumbnails.
	const [container, setContainer] = useState<HTMLDivElement | null>(null);
	const scrollContainerRef = useCallback((el: HTMLDivElement | null) => setContainer(el), []);

	useIsomorphicLayoutEffect(() => {
		view.setShape({ totalItems, itemHeight, overscan });
	}, [view, totalItems, itemHeight, overscan]);

	useIsomorphicLayoutEffect(() => {
		view.attach(container);
		return () => view.attach(null);
	}, [view, container]);

	useEffect(
		() => () => {
			deferredRows.destroy();
			view.destroy();
		},
		[deferredRows, view],
	);
	const syncRows = useCallback(
		(keys: readonly string[]) => deferredRows.sync(keys),
		[deferredRows],
	);

	const scrollToIndex = useCallback(
		(index: number, behavior: ScrollBehavior = 'smooth') => view.reveal(index, behavior),
		[view],
	);

	return {
		...(range ?? view.range),
		scrollContainerRef,
		scrollToIndex,
		syncRows,
		readyRows,
	};
}
