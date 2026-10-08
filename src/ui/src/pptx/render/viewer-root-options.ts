/**
 * The root display options every PowerPoint binding accepts: `initialSlide`,
 * `showToolbar` and `showThumbnails`. The defaults and the clamping live here
 * so React, Vue, Angular, Svelte and vanilla agree on them; each binding only
 * maps the resolved values onto its own template.
 *
 * - `showToolbar` (default `true`) gates the whole editor chrome around the
 *   canvas: the title bar, ribbon and toolbar, the quick-access toolbar, the
 *   protected-view and read-only banners, the mobile toolbar and the status
 *   bar. `false` leaves a bare canvas (plus the thumbnail pane, if shown).
 * - `showThumbnails` (default `true`) gates the slide thumbnail pane. It is a
 *   ceiling: the pane still follows the user's own collapse toggle and the
 *   `slidesPane` customisation panel switch.
 * - `initialSlide` (default `0`, zero-based) is the slide shown after each
 *   load, clamped into the deck with {@link resolveInitialSlideIndex}.
 */

/** The host-facing root display options, all optional. */
export interface ViewerRootOptions {
	/** Zero-based slide shown after load. Default `0`; clamped into the deck. */
	initialSlide?: number | undefined;
	/** Show the editor chrome (title bar, ribbon, toolbar, status bar). Default `true`. */
	showToolbar?: boolean | undefined;
	/** Show the slide thumbnail pane. Default `true`. */
	showThumbnails?: boolean | undefined;
}

/** {@link ViewerRootOptions} with every default applied. */
export interface ResolvedViewerRootOptions {
	initialSlide: number;
	showToolbar: boolean;
	showThumbnails: boolean;
}

/** The defaults the bindings document for the root display options. */
export const VIEWER_ROOT_OPTION_DEFAULTS: Readonly<ResolvedViewerRootOptions> = Object.freeze({
	initialSlide: 0,
	showToolbar: true,
	showThumbnails: true,
});

/**
 * Apply the defaults to the host's root display options. `initialSlide` is
 * only normalised here (a non-finite value falls back to `0`); clamp it into
 * a loaded deck with {@link resolveInitialSlideIndex}.
 */
export function resolveViewerRootOptions(
	options: ViewerRootOptions | undefined,
): ResolvedViewerRootOptions {
	const initialSlide = options?.initialSlide;
	return {
		initialSlide:
			typeof initialSlide === 'number' && Number.isFinite(initialSlide)
				? initialSlide
				: VIEWER_ROOT_OPTION_DEFAULTS.initialSlide,
		showToolbar: options?.showToolbar ?? VIEWER_ROOT_OPTION_DEFAULTS.showToolbar,
		showThumbnails: options?.showThumbnails ?? VIEWER_ROOT_OPTION_DEFAULTS.showThumbnails,
	};
}

/**
 * The slide to show after loading a deck of `slideCount` slides: the host's
 * `initialSlide` truncated to an integer and clamped into `[0, slideCount - 1]`.
 * An empty deck, a missing value and a non-finite value all give `0`.
 */
export function resolveInitialSlideIndex(
	initialSlide: number | undefined,
	slideCount: number,
): number {
	if (!(slideCount > 0)) {
		return 0;
	}
	const requested = resolveViewerRootOptions({ initialSlide }).initialSlide;
	return Math.min(Math.max(Math.trunc(requested), 0), Math.trunc(slideCount) - 1);
}
