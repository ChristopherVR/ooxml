/**
 * Keep thin strokes and borders at least ONE DEVICE PIXEL wide on screen.
 *
 * PowerPoint never paints a line thinner than one device pixel: a 0.1 pt and a
 * 1 pt line both come out as a solid, full-colour hairline at fit zoom. The
 * viewer lays a slide out at its authored size and then shrinks it with
 * `transform: scale(...)`, so a stroke that is already under a pixel gets
 * thinner still and the rasteriser anti-aliases it towards the background
 * (faint grey, or gone).
 *
 * The fix is split in two so it costs nothing where it does not apply:
 *
 *  - A slide STAGE (the element carrying the `scale()` transform) publishes
 *    {@link DEVICE_PX_VAR}: one device pixel expressed in the stage's own,
 *    unscaled slide pixels, `1px / (scale * devicePixelRatio)`. The ratio comes
 *    from {@link DEVICE_PIXEL_RATIO_VAR} on the root element, which
 *    {@link deviceStrokeStageStyle} keeps current (a `resolution` media query
 *    fires when the user zooms the browser or moves the window to another
 *    screen), so a binding does not have to re-render on a ratio change.
 *  - Every on-screen stroke / border width is written as
 *    `max(<authored>px, var(--pptx-device-px, 0px))` ({@link screenStrokeWidth}).
 *    Where the variable is unset (anything that is not inside a stage) the
 *    fallback is `0px` and the authored width is used verbatim; once the real
 *    width exceeds one device pixel `max()` picks it, so nothing changes.
 *
 * Export keeps authored widths: the capture clone pass
 * ({@link neutralizeDeviceStrokes}, run by `prepareExportClone`) and the
 * `foreignObject` computed-style read ({@link withAuthoredStrokeWidths}) pin
 * the variable to `0px`. The SVG exporter, print-to-SVG and saving never see
 * these styles at all: they work from the document model.
 *
 * Width `0` (and a non-finite width) stays `0`: a line the deck does not paint
 * must not be invented as a hairline.
 *
 * @module render/device-pixel-stroke
 */

import type { DeviceBox } from './device-pixel-exact';

/** One device pixel in the enclosing stage's unscaled slide pixels. */
export const DEVICE_PX_VAR = '--pptx-device-px';
/** `window.devicePixelRatio`, mirrored onto the root element as a plain number. */
export const DEVICE_PIXEL_RATIO_VAR = '--pptx-dpr';

/** Format a CSS pixel length without exponent notation (invalid in CSS). */
function cssPx(width: number): string {
	const rounded = Number(width.toFixed(4));
	return `${rounded > 0 ? rounded : 0.0001}px`;
}

/**
 * One device pixel as a CSS step that is never zero. `round()` with a zero step
 * is NaN, which would invalidate the whole declaration, so a missing variable
 * (outside a stage) and the exporters' `0px` pin both resolve to a 1/1000 px
 * step: authored widths survive to well below anything visible.
 */
const STEP = `max(var(${DEVICE_PX_VAR}, 0px), 0.001px)`;

/** The on-screen width {@link screenStrokeWidth} resolves a `px` literal to (CSS expression). */
function quantizedWidth(widthExpr: string): string {
	return `max(${STEP}, round(${widthExpr}, ${STEP}))`;
}

/**
 * The on-screen CSS width of a stroke or border authored `width` slide pixels
 * wide, the way PowerPoint paints it: a WHOLE number of device pixels, at least
 * one. PowerPoint's own exports (measured at 1280, 1366 and 1920 px wide) draw
 * a line `max(1, round(width_pt * px_per_pt))` pixels thick with no
 * anti-aliased fringe: 0.25 to 1 pt at 1280 px are all 1 px, 1.5 pt is 2 px,
 * 2 pt is 3 px; 2.25 pt at 1920 px is 5 px (ties round up). Outside a stage the
 * authored width is used as is, or `minOutsideStage` when that is larger. `'0px'` for a missing, zero, negative or
 * non-finite width.
 *
 * Must be written as a CSS PROPERTY (an inline `style`), not an SVG
 * presentation attribute: attributes cannot hold `var()`, `max()` or `round()`.
 * Keep the plain numeric attribute next to it for consumers that read
 * attributes.
 */
export function screenStrokeWidth(width: number | undefined, minOutsideStage = 0): string {
	if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) {
		return '0px';
	}
	const quantized = quantizedWidth(cssPx(width));
	return minOutsideStage > 0
		? `max(${quantized}, ${outsideStageFloor(minOutsideStage)})`
		: quantized;
}

/**
 * `min` px where no stage is publishing a device pixel (the variable is unset,
 * or an exporter pinned it to `0px`), nothing inside a stage, where the device
 * pixel itself is the floor. For lines that have always been at least one slide
 * pixel wide off screen (connectors): the 1000px / 0.001px thresholds sit far
 * outside any real device pixel.
 */
function outsideStageFloor(min: number): string {
	const dp = `var(${DEVICE_PX_VAR}, 1000px)`;
	const px = cssPx(min);
	return `max(clamp(0px, calc(${dp} - 100px), ${px}), clamp(0px, calc((0.001px - ${dp}) * 1000), ${px}))`;
}

/**
 * The shift (a CSS length) that puts a stroke of authored `width`, centred on
 * `coord` (slide px, along the axis the stroke is thin in), where PowerPoint's
 * export puts it: the stroke's near edge on the device pixel
 * `round(coord) - floor(n / 2)` for an `n` device pixel wide stroke, so a 1 px
 * line fills exactly one pixel row instead of straddling two. Add it as a
 * `translate` to an axis-aligned stroke only; diagonals stay anti-aliased in
 * PowerPoint too.
 *
 * `coord` is the stroke's position in the STAGE (the scaled element's own
 * coordinate space), which is where the device-pixel step is defined.
 */
export function snappedStrokeShift(coord: number, width: number | undefined): string {
	if (
		typeof width !== 'number' ||
		!Number.isFinite(width) ||
		width <= 0 ||
		!Number.isFinite(coord)
	) {
		return '0px';
	}
	const at = `${Number(coord.toFixed(4))}px`;
	const full = quantizedWidth(cssPx(width));
	return `calc(round(${at}, ${STEP}) - round(down, calc(${full} / 2), ${STEP}) + ${full} / 2 - ${at})`;
}

/**
 * The `transform` that snaps an axis-aligned stroke of `width` to the device
 * grid (see {@link snappedStrokeShift}): `axis` is the direction the stroke
 * RUNS in, so a horizontal line (`'h'`) shifts in y only, a vertical one
 * (`'v'`) in x only, and an outline whose edges all need it (`'box'`) in both.
 * `x`/`y` are the stroke's position in the stage.
 *
 * A `'box'` given its `size` snaps the far edges too: PowerPoint rounds every
 * edge on its own, so the right and bottom edges land on `round(x + w)` and
 * `round(y + h)`, not at the near edge's shift plus the authored size. That is
 * a translate to the near edge and a scale by `(w + farShift - nearShift) / w`:
 * a fraction of a device pixel over the whole rectangle, so the stroke width
 * stays within a fraction of a percent. Without a size only the top and left
 * edges are snapped.
 *
 * `x`, `y` and `size` are the element's box as the model lays it out (whole
 * pixels). `exact` is the box PowerPoint positions it at (see
 * {@link exactDeviceBox}); the snap rounds THAT, so a stroke lands on the same
 * device pixel as in PowerPoint instead of one off whenever the model's whole
 * pixel differs from the real position.
 */
export function snappedStrokeTransform(
	axis: 'h' | 'v' | 'box',
	x: number,
	y: number,
	width: number | undefined,
	size?: { width: number; height: number },
	exact?: DeviceBox,
	alongDashed = false,
): string | undefined {
	if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) {
		return undefined;
	}
	const ex = exact?.x ?? x;
	const ey = exact?.y ?? y;
	// A solid line is anti-aliased along its own axis; a DASHED one starts its
	// pattern on the snapped device pixel (a pixel centre for an odd pen), so a
	// hairline `sysDot` comes out as the uniform half-intensity line PowerPoint paints.
	const dx = axis === 'h' && !alongDashed ? '0px' : nearEdgeShift(x, ex, width);
	const dy = axis === 'v' && !alongDashed ? '0px' : nearEdgeShift(y, ey, width);
	const translate = `translate(${dx}, ${dy})`;
	if (axis !== 'box' || !size) {
		return translate;
	}
	const sx = farEdgeScale(ex, exact?.width ?? size.width, size.width, width);
	const sy = farEdgeScale(ey, exact?.height ?? size.height, size.height, width);
	return sx === '1' && sy === '1' ? translate : `${translate} scale(${sx}, ${sy})`;
}

/** The shift from the model's `coord` to where the stroke at `exact` snaps to. */
function nearEdgeShift(coord: number, exact: number, width: number): string {
	const shift = snappedStrokeShift(exact, width);
	const offset = Number((exact - coord).toFixed(4));
	return offset === 0 ? shift : `calc(${shift} + ${offset}px)`;
}

/**
 * The scale that moves a box's far edge, laid out `modelExtent` from its near
 * edge, onto the device pixel PowerPoint snaps `exact + exactExtent` to.
 */
function farEdgeScale(
	exact: number,
	exactExtent: number,
	modelExtent: number,
	width: number,
): string {
	if (
		!Number.isFinite(exact) ||
		!Number.isFinite(exactExtent) ||
		!Number.isFinite(modelExtent) ||
		modelExtent <= 0
	) {
		return '1';
	}
	const near = snappedStrokeShift(exact, width);
	const far = snappedStrokeShift(exact + exactExtent, width);
	return `calc((${Number(exactExtent.toFixed(4))}px + ${far} - ${near}) / ${Number(modelExtent.toFixed(4))}px)`;
}

/**
 * The on-screen CSS width of a `border-width`: never under one device pixel
 * inside a stage, the authored width anywhere else. Unlike
 * {@link screenStrokeWidth} it is NOT rounded to whole device pixels: Chromium
 * snaps a CSS border to whole pixels in the element's own (unscaled) layout
 * space, and a width rounded to a device pixel of a scaled stage (0.94 px at a
 * scale of 1.0667, say) lands under one layout pixel there and can drop out or
 * move a row, where `max()` leaves the snap to the browser.
 */
export function screenBorderWidth(width: number | undefined): string {
	if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) {
		return '0px';
	}
	return `max(${cssPx(width)}, var(${DEVICE_PX_VAR}, 0px))`;
}

/**
 * A CSS `border` shorthand whose width is {@link screenBorderWidth}; a missing
 * `style` is `solid`.
 */
export function screenBorder(width: number, style: string | undefined, color: string): string {
	return `${screenBorderWidth(width)} ${style ?? 'solid'} ${color}`;
}

interface DevicePixelRatioWindow {
	devicePixelRatio?: number;
	document?: { documentElement?: { style?: { setProperty(name: string, value: string): void } } };
	matchMedia?: (query: string) => {
		addEventListener?: (type: 'change', listener: () => void, options?: { once?: boolean }) => void;
	};
}

const watchedWindows = new WeakSet<DevicePixelRatioWindow>();

/**
 * Mirror `devicePixelRatio` onto the root element as {@link DEVICE_PIXEL_RATIO_VAR}
 * and keep it current. Idempotent per window; a no-op outside a browser.
 */
export function watchDevicePixelRatio(
	win: DevicePixelRatioWindow | undefined = typeof window === 'undefined'
		? undefined
		: (window as unknown as DevicePixelRatioWindow),
): void {
	if (!win || watchedWindows.has(win)) {
		return;
	}
	const style = win.document?.documentElement?.style;
	if (!style) {
		return;
	}
	watchedWindows.add(win);
	const publish = (): void => {
		const ratio = win.devicePixelRatio;
		const safe = typeof ratio === 'number' && Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
		style.setProperty(DEVICE_PIXEL_RATIO_VAR, String(safe));
		// A `resolution` query only matches the CURRENT ratio, so re-arm it for
		// the new one each time it stops matching.
		win.matchMedia?.(`(resolution: ${safe}dppx)`).addEventListener?.('change', publish, {
			once: true,
		});
	};
	publish();
}

/**
 * The custom property a slide stage scaled by `scale` publishes so the strokes
 * inside it can resolve one device pixel (see the module comment). Also makes
 * sure the root carries the live device-pixel ratio.
 *
 * Returns an empty map for a non-positive or non-finite scale, which leaves
 * authored widths in place.
 */
export function deviceStrokeStageStyle(scale: number): Record<string, string> {
	if (!Number.isFinite(scale) || scale <= 0) {
		return {};
	}
	watchDevicePixelRatio();
	return {
		[DEVICE_PX_VAR]: `calc(1px / (${Number(scale.toFixed(6))} * var(${DEVICE_PIXEL_RATIO_VAR}, 1)))`,
	};
}

type StyledElement = Element & ElementCSSInlineStyle;

/** Elements that DEFINE the variable (a user writes `var(--pptx-device-px`, not `--pptx-device-px:`). */
const CARRIER_SELECTOR = `[style*="${DEVICE_PX_VAR}:"]`;

/**
 * The elements whose variable decides what `root`'s subtree paints: `root`
 * itself when it, or an ancestor, is inside a stage, plus every nested stage.
 * A subtree with no stage anywhere is already at authored widths and is left
 * byte-for-byte alone.
 */
function carriers(root: StyledElement): StyledElement[] {
	const nested = [...root.querySelectorAll<StyledElement>(CARRIER_SELECTOR)];
	return root.closest(CARRIER_SELECTOR) ? [root, ...nested.filter((el) => el !== root)] : nested;
}

/**
 * Pin {@link DEVICE_PX_VAR} to `0px` on `root` and on every nested stage, so
 * the subtree paints authored widths. For a capture clone, where nothing needs
 * restoring (html2canvas clones the whole document, so a stage ANCESTOR of the
 * captured element is still there and is overridden at `root`).
 */
export function neutralizeDeviceStrokes(root: StyledElement): void {
	for (const el of carriers(root)) {
		el.style?.setProperty(DEVICE_PX_VAR, '0px');
	}
}

/**
 * Run `read` with the LIVE subtree temporarily at authored widths, then put
 * {@link DEVICE_PX_VAR} back exactly as it was on every touched element. For
 * export paths that copy computed styles from the live tree; the swap is
 * synchronous, so no frame is painted in between.
 *
 * Only the variable is saved and restored, never the whole `style`
 * attribute: the attribute is the browser's serialisation of the inline
 * declarations, which rounds numbers (Chromium writes a framework-set
 * `scale(0.9958333)` back as `scale(0.995833)`). Re-parsing it would shift
 * the live stage's transform by a fraction of a pixel for good, changing the
 * measured size, and so the pixel size, of every later export.
 */
export function withAuthoredStrokeWidths<T>(root: StyledElement, read: () => T): T {
	const saved = carriers(root).map((el) => ({
		el,
		hadStyle: el.hasAttribute('style'),
		value: el.style?.getPropertyValue(DEVICE_PX_VAR) ?? '',
		priority: el.style?.getPropertyPriority(DEVICE_PX_VAR) ?? '',
	}));
	for (const { el } of saved) {
		el.style?.setProperty(DEVICE_PX_VAR, '0px');
	}
	try {
		return read();
	} finally {
		for (const { el, hadStyle, value, priority } of saved) {
			if (value) {
				el.style?.setProperty(DEVICE_PX_VAR, value, priority);
			} else {
				el.style?.removeProperty(DEVICE_PX_VAR);
			}
			if (!hadStyle && !el.getAttribute('style')) {
				el.removeAttribute('style');
			}
		}
	}
}
