/**
 * `slide-thumbnail-badge` - the transition badge a slide thumbnail draws in its
 * top-right corner, shared by every binding's `SlideThumbnail`.
 *
 * WHY this lives in shared: whether a slide shows the badge, what it says and
 * how it is drawn are plain decisions with no framework in them. The badge
 * first existed only in React's thumbnail, so the other bindings' thumbnails
 * drew none; one descriptor and one style keep the five identical.
 *
 * @module render/slide-thumbnail-badge
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import type { SlideThumbnailLayout } from './editor-chrome/metrics';
import { SLIDE_TRANSITION_OPTIONS } from './slide-transition-options';

/** A binding's translator: an i18n key and its `{{name}}` parameters to text. */
export type SlideThumbnailTranslate = (key: string, params?: Record<string, string>) => string;

/** What a thumbnail's transition badge shows. */
export interface SlideThumbnailTransitionBadge {
	/** The transition's name, the badge's text. */
	label: string;
	/** The badge's tooltip ("Transition: Fade"). */
	title: string;
}

/** i18n key of the badge tooltip; `{{name}}` is the transition's name. */
export const SLIDE_THUMBNAIL_TRANSITION_TOOLTIP_KEY = 'pptx.slideThumbnail.transitionTooltip';

/**
 * The transition badge of `slide`'s thumbnail, or `undefined` when it has none.
 * A slide without a transition, or with `none` or `cut` (an instant change,
 * nothing to play), shows no badge. The label is the transition type's
 * translated name, or the raw type when the catalogue does not list it.
 */
export function slideThumbnailTransitionBadge(
	slide: Pick<PptxSlide, 'transition'>,
	translate: SlideThumbnailTranslate,
): SlideThumbnailTransitionBadge | undefined {
	const type = slide.transition?.type;
	if (!type || type === 'none' || type === 'cut') {
		return undefined;
	}
	const option = SLIDE_TRANSITION_OPTIONS.find((entry) => entry.value === type);
	const label = option ? translate(option.i18nKey) : type;
	return { label, title: translate(SLIDE_THUMBNAIL_TRANSITION_TOOLTIP_KEY, { name: label }) };
}

/**
 * Inline style of the transition badge, placed in the thumbnail box after the
 * slide stage. `layout` is the thumbnail's `slideThumbnailLayout`: the badge is
 * drawn at its stage scale, 2px in from the drawn slide's top-right corner in
 * slide pixels, so it keeps its size relative to the slide at any thumbnail
 * width. The corner is the drawn slide's (`layout.width` from the box's left
 * edge), not the box's, which a container can make wider or narrower than the
 * slide. It sits above every element and takes no pointer events.
 */
export function slideThumbnailTransitionBadgeStyle(
	layout: Pick<SlideThumbnailLayout, 'width' | 'scale'>,
): Record<string, string> {
	const { scale } = layout;
	const inset = 2 * scale;
	return {
		position: 'absolute',
		top: `${String(inset)}px`,
		right: `calc(100% - ${String(layout.width - inset)}px)`,
		zIndex: '10',
		transform: `scale(${String(scale)})`,
		transformOrigin: 'top right',
		padding: '1px 4px',
		borderRadius: '4px',
		background:
			'color-mix(in oklab, var(--pptx-primary, var(--office-accent, #2563eb)) 80%, transparent)',
		color: '#fff',
		fontSize: '7px',
		lineHeight: '1.25',
		whiteSpace: 'nowrap',
		pointerEvents: 'none',
	};
}
