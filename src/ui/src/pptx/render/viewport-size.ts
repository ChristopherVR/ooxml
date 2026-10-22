/**
 * The size a viewer viewport offers its fitted slide, measured without whole-pixel rounding.
 *
 * `clientWidth` / `clientHeight` are snapped to whole CSS pixels, so a viewport laid out at
 * 801.5px reports 802 and a slide fitted to it is half a pixel wider than the box that holds
 * it. Every binding measures its viewport through this module (each wraps
 * {@link observeViewportClientSize} or {@link measureViewportClientSize} in its own reactive
 * primitive), so the measurement rule cannot drift between them.
 *
 * The exact size is only taken where it can be trusted: in a plain context, where nothing
 * transforms or CSS-zooms the viewport, `getBoundingClientRect()` is the laid-out border box in
 * the viewport's own CSS pixels (also under browser zoom). A CSS `transform` or `zoom` on the
 * viewport or an ancestor scales the rect, while the computed `width` / `height` stay local, so
 * the two disagree; the slide inside is then laid out and painted through that scale, where its
 * own layout rounding can move its edges by a fraction of a pixel. There a whole pixel is
 * reserved instead, as for a scrollbar.
 */

/** The layout metrics a viewport element reports, one value per axis. */
export interface ViewportBoxMetrics {
	/** `clientWidth` / `clientHeight`: padding box without scrollbars, snapped to whole pixels. */
	clientWidth: number;
	clientHeight: number;
	/** `offsetWidth` / `offsetHeight`: border box, snapped the same way. */
	offsetWidth: number;
	offsetHeight: number;
	/** `getBoundingClientRect()` size: the border box before snapping (and after transforms). */
	rectWidth: number;
	rectHeight: number;
	/**
	 * Used border widths from `getComputedStyle`, which keep their fraction: under browser zoom
	 * a `1px` border is 1.111px at 90%. Without them the borders are taken to be whole pixels.
	 */
	borderLeft?: number;
	borderRight?: number;
	borderTop?: number;
	borderBottom?: number;
	/**
	 * Used `width` / `height` from `getComputedStyle`, in local CSS pixels: the border box under
	 * `box-sizing: border-box`, the content box otherwise. Missing or non-finite (`auto` on a box
	 * without layout, a test DOM) falls back to the rect.
	 */
	cssWidth?: number;
	cssHeight?: number;
	/** `box-sizing` from `getComputedStyle`; anything else counts as unknown. */
	boxSizing?: string;
	/** Used padding from `getComputedStyle`, needed to turn a content-box size into a border box. */
	paddingLeft?: number;
	paddingRight?: number;
	paddingTop?: number;
	paddingBottom?: number;
	/**
	 * `element.currentCSSZoom` where the browser has it: the product of the CSS `zoom` of the
	 * viewport and its ancestors (browser zoom is not part of it). Anything but 1 means the
	 * context is zoomed.
	 */
	cssZoom?: number;
}

export interface ViewportClientSize {
	width: number;
	height: number;
}

/** A layout unit, the finest step a browser lays boxes out in (Blink and WebKit use 1/64px). */
const LAYOUT_UNIT = 1 / 64;
/**
 * Padding is laid out rounded down to a layout unit of the zoomed page, which is coarser in
 * local pixels below 100% browser zoom, so a content-box viewport's rect may be smaller than
 * its computed size plus padding by this much per padded side (browser zoom goes down to 25%).
 */
const PADDING_SNAP = LAYOUT_UNIT / 0.25;

function finiteOrZero(value: number): number {
	return Number.isFinite(value) && value > 0 ? value : 0;
}

function isFiniteNumber(value: number | undefined): value is number {
	return value !== undefined && Number.isFinite(value);
}

/** Half a unit in the sixth significant digit: how far a serialized computed length can be off. */
function serializationError(value: number): number {
	const magnitude = Math.abs(value);
	return magnitude === 0 ? 0 : 0.5 * 10 ** (Math.floor(Math.log10(magnitude)) - 5);
}

/** The local border box from the computed size, and the range a plain context's rect is in. */
interface LocalBorderBox {
	size: number;
	/** Lowest rect a plain context lays out for it (serialization and padding rounding). */
	low: number;
	/** Highest such rect (serialization only: layout never rounds padding up). */
	high: number;
	/** Never more than the laid-out border box: `size` less its serialization and padding rounding. */
	floor: number;
	/**
	 * Whether `size` is the laid-out size up to serialization. With content-box padding, which is
	 * laid out rounded down, it is not: a rect between `floor` and `size` is then either that
	 * rounding or a slight scale, and cannot be told apart.
	 */
	exact: boolean;
}

function localBorderBox(
	cssSize: number | undefined,
	boxSizing: string | undefined,
	borders: number,
	paddingStart: number | undefined,
	paddingEnd: number | undefined,
): LocalBorderBox | undefined {
	if (!isFiniteNumber(cssSize) || cssSize < 0) {
		return undefined;
	}
	if (boxSizing === 'border-box') {
		const error = serializationError(cssSize);
		return {
			size: cssSize,
			low: cssSize - error,
			high: cssSize + error,
			floor: cssSize - error,
			exact: true,
		};
	}
	if (boxSizing !== 'content-box' || !isFiniteNumber(paddingStart) || !isFiniteNumber(paddingEnd)) {
		return undefined;
	}
	const start = Math.max(0, paddingStart);
	const end = Math.max(0, paddingEnd);
	const size = cssSize + start + end + borders;
	const error =
		serializationError(cssSize) +
		serializationError(start) +
		serializationError(end) +
		serializationError(borders);
	const snap = (start > 0 ? PADDING_SNAP : 0) + (end > 0 ? PADDING_SNAP : 0);
	return {
		size,
		low: size - error - snap,
		high: size + error,
		floor: size - error - snap,
		exact: snap === 0,
	};
}

/**
 * One axis. A scrollbar is whatever more than a pixel beyond the borders separates the snapped
 * border and client boxes; with one, the snapped client box less a pixel is used, so the result
 * never exceeds the box.
 *
 * Otherwise the context decides. In a plain one (no CSS zoom, and a rect that matches the local
 * computed border box to within its serialization and padding rounding) the rect less the used
 * borders is the exact padding box. Past ten thousand pixels a serialized size is only good to
 * a twentieth of a pixel, so the rect is capped to a layout unit above the computed size there.
 * In a transformed or zoomed context the smaller of the snapped client size and the computed
 * padding box, less a pixel, is used: never more than the box, whatever the slide's own layout
 * rounding does under the scale.
 *
 * Without computed sizes (test DOMs), the rect is used as before: a rect that differs from the
 * border box by a pixel or more is scaled, so one pixel is reserved; without border widths the
 * borders are taken to be whole pixels, so the client box carries the border box's snapping
 * error.
 */
function fractionalAxis(
	client: number,
	offset: number,
	rect: number,
	borderStart: number | undefined,
	borderEnd: number | undefined,
	cssSize: number | undefined,
	boxSizing: string | undefined,
	paddingStart: number | undefined,
	paddingEnd: number | undefined,
	zoomed: boolean,
): number {
	const snapped = finiteOrZero(client);
	// Not laid out (or a DOM without layout): nothing to correct.
	if (snapped === 0 || !(finiteOrZero(offset) > 0)) {
		return snapped;
	}
	const bordersKnown = isFiniteNumber(borderStart) && isFiniteNumber(borderEnd);
	const borders = bordersKnown ? Math.max(0, borderStart) + Math.max(0, borderEnd) : 0;
	// Snapping the border and client boxes adds under half a pixel each, so less than a pixel
	// between them is rounding; a scrollbar is wider.
	if (bordersKnown && offset - snapped - borders >= 1) {
		return Math.max(0, snapped - 1);
	}
	const local = bordersKnown
		? localBorderBox(cssSize, boxSizing, borders, paddingStart, paddingEnd)
		: undefined;
	if (local) {
		const plain = !zoomed && Number.isFinite(rect) && rect >= local.low && rect <= local.high;
		if (plain) {
			// Where padding rounding leaves the size uncertain, its lower bound is used, so a rect a
			// slight scale has grown cannot pass for the laid-out size.
			const size = local.exact ? Math.min(rect, local.floor + LAYOUT_UNIT) : local.floor;
			return Math.max(0, size - borders);
		}
		return Math.max(0, Math.min(snapped, local.size - borders) - 1);
	}
	if (!Number.isFinite(rect)) {
		return snapped;
	}
	const error = rect - offset;
	if (zoomed || Math.abs(error) >= 1) {
		return Math.max(0, snapped - 1);
	}
	if (!bordersKnown) {
		return Math.max(0, snapped + error);
	}
	return Math.max(0, rect - borders);
}

/** A CSS zoom other than 1 on the viewport or an ancestor (`currentCSSZoom` is a float). */
function isZoomed(cssZoom: number | undefined): boolean {
	return isFiniteNumber(cssZoom) && Math.abs(cssZoom - 1) > 1e-6;
}

/**
 * The fractional client size of a viewport from its snapped and unsnapped metrics: what
 * `clientWidth` / `clientHeight` would report without rounding, in the viewport's local CSS
 * pixels, never more than the box actually holds. DOM-free so it is unit-testable;
 * {@link measureViewportClientSize} reads the metrics from an element.
 */
export function fractionalClientSize(metrics: ViewportBoxMetrics): ViewportClientSize {
	const zoomed = isZoomed(metrics.cssZoom);
	return {
		width: fractionalAxis(
			metrics.clientWidth,
			metrics.offsetWidth,
			metrics.rectWidth,
			metrics.borderLeft,
			metrics.borderRight,
			metrics.cssWidth,
			metrics.boxSizing,
			metrics.paddingLeft,
			metrics.paddingRight,
			zoomed,
		),
		height: fractionalAxis(
			metrics.clientHeight,
			metrics.offsetHeight,
			metrics.rectHeight,
			metrics.borderTop,
			metrics.borderBottom,
			metrics.cssHeight,
			metrics.boxSizing,
			metrics.paddingTop,
			metrics.paddingBottom,
			zoomed,
		),
	};
}

function pixels(value: string | undefined): number | undefined {
	// `auto`, an empty string or a non-pixel value parse to NaN or a non-finite number.
	const parsed =
		value === undefined || !value.endsWith('px') ? Number.NaN : Number.parseFloat(value);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** Measure `element`'s fractional client size now (a synchronous layout read). */
export function measureViewportClientSize(element: HTMLElement): ViewportClientSize {
	const rect = element.getBoundingClientRect();
	const view = element.ownerDocument?.defaultView;
	const style = view?.getComputedStyle ? view.getComputedStyle(element) : undefined;
	const { currentCSSZoom } = element as { currentCSSZoom?: number };
	return fractionalClientSize({
		clientWidth: element.clientWidth,
		clientHeight: element.clientHeight,
		offsetWidth: element.offsetWidth,
		offsetHeight: element.offsetHeight,
		rectWidth: rect.width,
		rectHeight: rect.height,
		borderLeft: pixels(style?.borderLeftWidth),
		borderRight: pixels(style?.borderRightWidth),
		borderTop: pixels(style?.borderTopWidth),
		borderBottom: pixels(style?.borderBottomWidth),
		cssWidth: pixels(style?.width),
		cssHeight: pixels(style?.height),
		boxSizing: style?.boxSizing,
		paddingLeft: pixels(style?.paddingLeft),
		paddingRight: pixels(style?.paddingRight),
		paddingTop: pixels(style?.paddingTop),
		paddingBottom: pixels(style?.paddingBottom),
		cssZoom: typeof currentCSSZoom === 'number' ? currentCSSZoom : undefined,
	});
}

/**
 * Report `element`'s fractional client size now, synchronously, and again whenever it changes.
 *
 * The first report is made before this returns, so a binding that calls it before the browser
 * paints (a layout effect, a mount hook) never paints an unfitted slide. Returns a disposer.
 * Falls back to the single initial measurement where `ResizeObserver` does not exist (SSR,
 * some test DOMs).
 *
 * The client size changes with the content box (a border-box viewport's padding, a scrollbar)
 * and with the border box (a content-box viewport's padding), and a `ResizeObserver` watches
 * only one box, so both are observed. A size equal to the last report is not reported again,
 * so a resize both observers see is reported once.
 */
export function observeViewportClientSize(
	element: HTMLElement,
	onSize: (size: ViewportClientSize) => void,
): () => void {
	let last = measureViewportClientSize(element);
	onSize(last);
	if (typeof ResizeObserver === 'undefined') {
		return () => undefined;
	}
	const remeasure = () => {
		const next = measureViewportClientSize(element);
		if (next.width === last.width && next.height === last.height) {
			return;
		}
		last = next;
		onSize(next);
	};
	const observers = (['content-box', 'border-box'] as const).map((box) => {
		const observer = new ResizeObserver(remeasure);
		observer.observe(element, { box });
		return observer;
	});
	return () => {
		for (const observer of observers) {
			observer.disconnect();
		}
	};
}
