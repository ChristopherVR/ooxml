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
 * The on-screen CSS width of a stroke or border authored `width` slide pixels
 * wide: never under one device pixel inside a stage, the authored width
 * anywhere else. `'0px'` for a missing, zero, negative or non-finite width.
 *
 * Must be written as a CSS PROPERTY (an inline `style`), not an SVG
 * presentation attribute: attributes cannot hold `var()` or `max()`. Keep the
 * plain numeric attribute next to it for consumers that read attributes.
 */
export function screenStrokeWidth(width: number | undefined): string {
	if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) {
		return '0px';
	}
	return `max(${cssPx(width)}, var(${DEVICE_PX_VAR}, 0px))`;
}

/**
 * A CSS `border` shorthand whose width is {@link screenStrokeWidth}; a missing
 * `style` is `solid`.
 */
export function screenBorder(width: number, style: string | undefined, color: string): string {
	return `${screenStrokeWidth(width)} ${style ?? 'solid'} ${color}`;
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
 * every touched `style` attribute back exactly as it was. For export paths
 * that copy computed styles from the live tree; the swap is synchronous, so no
 * frame is painted in between.
 */
export function withAuthoredStrokeWidths<T>(root: StyledElement, read: () => T): T {
	const saved = carriers(root).map((el) => ({ el, style: el.getAttribute('style') }));
	for (const { el } of saved) {
		el.style?.setProperty(DEVICE_PX_VAR, '0px');
	}
	try {
		return read();
	} finally {
		for (const { el, style } of saved) {
			if (style === null) {
				el.removeAttribute('style');
			} else {
				el.setAttribute('style', style);
			}
		}
	}
}
