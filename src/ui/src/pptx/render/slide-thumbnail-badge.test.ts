import { describe, expect, it } from 'vitest';

import {
	SLIDE_THUMBNAIL_TRANSITION_TOOLTIP_KEY,
	slideThumbnailTransitionBadge,
	slideThumbnailTransitionBadgeStyle,
} from './slide-thumbnail-badge';

const translate = (key: string, params?: Record<string, string>): string =>
	params ? `${key}(${JSON.stringify(params)})` : key;

describe('slideThumbnailTransitionBadge', () => {
	it('shows no badge without a transition, or for none and cut', () => {
		expect(slideThumbnailTransitionBadge({}, translate)).toBeUndefined();
		expect(
			slideThumbnailTransitionBadge({ transition: { type: 'none' } }, translate),
		).toBeUndefined();
		expect(
			slideThumbnailTransitionBadge({ transition: { type: 'cut' } }, translate),
		).toBeUndefined();
	});

	it('labels the badge with the translated transition name', () => {
		expect(
			slideThumbnailTransitionBadge({ transition: { type: 'fade' } }, translate),
		).toStrictEqual({
			label: 'pptx.ribbon.transition.fade',
			title: `${SLIDE_THUMBNAIL_TRANSITION_TOOLTIP_KEY}({"name":"pptx.ribbon.transition.fade"})`,
		});
	});

	it('falls back to the raw type for a type the catalogue does not list', () => {
		const badge = slideThumbnailTransitionBadge(
			{ transition: { type: 'notInCatalogue' as 'fade' } },
			translate,
		);
		expect(badge?.label).toBe('notInCatalogue');
	});
});

describe('slideThumbnailTransitionBadgeStyle', () => {
	it('draws the badge at the slide scale, inset from the top-right corner', () => {
		const style = slideThumbnailTransitionBadgeStyle({ width: 240, scale: 0.25 });
		expect(style).toMatchObject({
			position: 'absolute',
			top: '0.5px',
			right: 'calc(100% - 239.5px)',
			transform: 'scale(0.25)',
			transformOrigin: 'top right',
			pointerEvents: 'none',
		});
	});

	// The box fills its container, which can be wider or narrower than the drawn
	// slide when a caller passes no matching width. `right` is measured from the
	// box's right edge, so it subtracts the drawn width from the box's: in a 188px
	// box a 156px slide puts the badge 32.325px in, on the slide's corner.
	it('anchors to the drawn slide corner, not the box corner', () => {
		const style = slideThumbnailTransitionBadgeStyle({ width: 156, scale: 0.1625 });
		expect(style.top).toBe('0.325px');
		expect(style.right).toBe('calc(100% - 155.675px)');
	});
});
