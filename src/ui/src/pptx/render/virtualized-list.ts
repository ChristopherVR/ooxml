/** Framework-neutral virtual-list range used by every thumbnail sidebar. */
export interface VirtualizedRange {
	startIndex: number;
	endIndex: number;
	totalHeight: number;
	offsetY: number;
	visibleRange: { start: number; end: number };
}

export const DEFAULT_VIRTUAL_OVERSCAN = 5;
export const SLIDE_VIRTUALIZATION_THRESHOLD = 50;
/** Viewport height assumed before the scroll container has been laid out. */
export const DEFAULT_VIRTUAL_VIEWPORT_HEIGHT = 600;

/** Compute the rendered item window for fixed-height virtual lists. */
export function computeVirtualRange(
	totalItems: number,
	itemHeight: number,
	scrollTop: number,
	viewportHeight: number,
	overscan = DEFAULT_VIRTUAL_OVERSCAN,
): VirtualizedRange {
	const safeCount = Math.max(0, totalItems);
	const safeItemHeight = Math.max(itemHeight, 1);
	const totalHeight = safeCount * safeItemHeight;
	if (safeCount === 0) {
		return {
			startIndex: 0,
			endIndex: -1,
			totalHeight: 0,
			offsetY: 0,
			visibleRange: { start: 0, end: -1 },
		};
	}

	const visibleStart = Math.min(safeCount - 1, Math.max(0, Math.floor(scrollTop / safeItemHeight)));
	const visibleEnd = Math.min(
		safeCount - 1,
		Math.max(visibleStart, Math.floor((scrollTop + Math.max(0, viewportHeight)) / safeItemHeight)),
	);
	const safeOverscan = Math.max(0, overscan);
	const startIndex = Math.max(0, visibleStart - safeOverscan);
	const endIndex = Math.min(safeCount - 1, visibleEnd + safeOverscan);
	return {
		startIndex,
		endIndex,
		totalHeight,
		offsetY: startIndex * safeItemHeight,
		visibleRange: { start: visibleStart, end: visibleEnd },
	};
}

/** True when two ranges render the same rows at the same place. */
export function sameRenderedWindow(a: VirtualizedRange, b: VirtualizedRange): boolean {
	return (
		a.startIndex === b.startIndex &&
		a.endIndex === b.endIndex &&
		a.totalHeight === b.totalHeight &&
		a.offsetY === b.offsetY
	);
}

/** The list a virtual window lays out: `totalItems` rows of `itemHeight` pixels. */
export interface VirtualListShape {
	totalItems: number;
	itemHeight: number;
	overscan?: number;
}

/** Frames, timers and the clock a virtual window runs on; replaceable in tests. */
export interface FrameScheduler {
	/** Run `callback` before the next paint. */
	request(callback: () => void): unknown;
	cancel(handle: unknown): void;
	/** Run `callback` after `ms` milliseconds; defaults to `setTimeout`. */
	delay?(callback: () => void, ms: number): unknown;
	cancelDelay?(handle: unknown): void;
	/** Milliseconds on a monotonic clock; defaults to `performance.now`. */
	now?(): number;
}

/** Deferred row fills wait this long after the last scroll event. */
export const DEFERRED_FILL_SETTLE_MS = 120;
/** Deferred row fills stop for the frame once they have taken this long. */
export const DEFERRED_FILL_FRAME_BUDGET_MS = 8;

const animationFrames: FrameScheduler = {
	request: (callback) =>
		typeof requestAnimationFrame === 'function'
			? requestAnimationFrame(() => callback())
			: setTimeout(callback, 16),
	cancel: (handle) => {
		if (typeof cancelAnimationFrame === 'function') {
			cancelAnimationFrame(handle as number);
		} else {
			clearTimeout(handle as ReturnType<typeof setTimeout>);
		}
	},
};

/**
 * The scroll window of one virtual list. Every binding's thumbnail rail holds
 * one: scroll and resize events are read at most once per frame, and
 * `onChange` fires only when the rendered rows change, so a rail re-renders
 * when a row enters or leaves, not on every scroll event.
 */
export interface VirtualWindow {
	/** The window last computed. */
	readonly range: VirtualizedRange;
	/** Track a scroll container; `null` stops tracking the current one. */
	attach(container: HTMLElement | null): void;
	/** Lay out a new item count or height; recomputes now. */
	setShape(shape: VirtualListShape): void;
	/**
	 * Scroll the container the least distance that shows row `index` whole,
	 * then recompute now, so the caller renders the window around that row in
	 * the same update.
	 */
	reveal(index: number, behavior?: ScrollBehavior): void;
	/**
	 * Run `fill` for row `key` once scrolling has paused, a frame's budget at a
	 * time, so a fast scroll paints row shells at frame rate and fills the rows
	 * still in view when it stops. A new request for a key replaces its pending
	 * one. A fill that returns `true` only schedules its work (a framework
	 * renders the row later), so it ends the frame's fills: the budget cannot
	 * see that work.
	 */
	defer(key: string, fill: () => boolean | void): void;
	/** Drop the pending fill of a row that left the window. */
	cancelDeferred(key: string): void;
	/** Stop listening and drop any pending frame and fill. */
	destroy(): void;
}

/** Create the scroll window for a fixed-height virtual list. */
export function createVirtualWindow(
	shape: VirtualListShape,
	onChange: (range: VirtualizedRange) => void,
	frames: FrameScheduler = animationFrames,
): VirtualWindow {
	let current = shape;
	let container: HTMLElement | null = null;
	let resizeObserver: ResizeObserver | null = null;
	let pending: unknown = null;

	const computeNow = (): VirtualizedRange =>
		computeVirtualRange(
			current.totalItems,
			current.itemHeight,
			container?.scrollTop ?? 0,
			container?.clientHeight || DEFAULT_VIRTUAL_VIEWPORT_HEIGHT,
			current.overscan,
		);
	let range = computeNow();

	const cancelPending = (): void => {
		if (pending !== null) {
			frames.cancel(pending);
			pending = null;
		}
	};

	const recompute = (): void => {
		cancelPending();
		const next = computeNow();
		const changed = !sameRenderedWindow(next, range);
		range = next;
		if (changed) {
			onChange(next);
		}
	};

	const schedule = (): void => {
		if (pending === null) {
			pending = frames.request(() => {
				pending = null;
				recompute();
			});
		}
	};

	// ── Deferred row fills ──
	const now = (): number => frames.now?.() ?? performance.now();
	const fills = new Map<string, () => boolean | void>();
	let lastScrollAt = Number.NEGATIVE_INFINITY;
	let fillFrame: unknown = null;
	let fillTimer: unknown = null;
	const delay = (callback: () => void, ms: number): unknown =>
		frames.delay ? frames.delay(callback, ms) : setTimeout(callback, ms);
	const cancelFillWork = (): void => {
		if (fillFrame !== null) {
			frames.cancel(fillFrame);
			fillFrame = null;
		}
		if (fillTimer !== null) {
			if (frames.cancelDelay) {
				frames.cancelDelay(fillTimer);
			} else {
				clearTimeout(fillTimer as ReturnType<typeof setTimeout>);
			}
			fillTimer = null;
		}
	};
	const pumpFills = (): void => {
		fillFrame = null;
		const quietFor = now() - lastScrollAt;
		if (quietFor < DEFERRED_FILL_SETTLE_MS) {
			fillTimer = delay(() => {
				fillTimer = null;
				scheduleFills();
			}, DEFERRED_FILL_SETTLE_MS - quietFor);
			return;
		}
		const start = now();
		for (const [key, fill] of fills) {
			fills.delete(key);
			if (fill() === true || now() - start >= DEFERRED_FILL_FRAME_BUDGET_MS) {
				break;
			}
		}
		if (fills.size > 0) {
			scheduleFills();
		}
	};
	const scheduleFills = (): void => {
		if (fillFrame === null && fillTimer === null && fills.size > 0) {
			fillFrame = frames.request(pumpFills);
		}
	};
	const onScroll = (): void => {
		lastScrollAt = now();
		schedule();
	};

	let anchorBefore = '';
	const detach = (): void => {
		if (container) {
			container.style.overflowAnchor = anchorBefore;
		}
		container?.removeEventListener('scroll', onScroll);
		resizeObserver?.disconnect();
		resizeObserver = null;
		container = null;
	};

	return {
		get range() {
			return range;
		},
		attach(el) {
			if (el === container) {
				return;
			}
			detach();
			container = el;
			if (el) {
				// The window swaps rows above the viewport as it scrolls; scroll
				// anchoring would then undo a reveal's scroll.
				anchorBefore = el.style.overflowAnchor;
				el.style.overflowAnchor = 'none';
				el.addEventListener('scroll', onScroll, { passive: true });
				if (typeof ResizeObserver !== 'undefined') {
					resizeObserver = new ResizeObserver(schedule);
					resizeObserver.observe(el);
				}
			}
			recompute();
		},
		setShape(next) {
			current = next;
			recompute();
		},
		reveal(index, behavior) {
			const el = container;
			if (el && index >= 0 && index < current.totalItems) {
				const itemHeight = Math.max(current.itemHeight, 1);
				const viewport = el.clientHeight || DEFAULT_VIRTUAL_VIEWPORT_HEIGHT;
				const top = index * itemHeight;
				const bottom = top + itemHeight;
				const target =
					top < el.scrollTop
						? top
						: bottom > el.scrollTop + viewport
							? Math.max(0, bottom - viewport)
							: null;
				if (target !== null && behavior === 'smooth' && typeof el.scrollTo === 'function') {
					el.scrollTo({ top: target, behavior });
				} else if (target !== null) {
					el.scrollTop = target;
				}
			}
			recompute();
		},
		defer(key, fill) {
			fills.delete(key);
			fills.set(key, fill);
			scheduleFills();
		},
		cancelDeferred(key) {
			fills.delete(key);
		},
		destroy() {
			cancelPending();
			cancelFillWork();
			fills.clear();
			detach();
		},
	};
}
