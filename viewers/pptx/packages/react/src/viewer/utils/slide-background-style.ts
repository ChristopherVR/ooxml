import type { SlideBackgroundSize } from 'ooxml-ui/pptx';
import { getSlideBackgroundStyle } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import type { CSSProperties } from 'react';

/**
 * Adapt the framework-neutral kebab-case slide background map for React.
 *
 * `slideSize` (the deck's authored px size) is only needed to anchor a
 * `shadeToTitle` gradient on its title placeholder; omitting it keeps the
 * plain authored gradient (see `getSlideBackgroundStyle`).
 */
export function getReactSlideBackgroundStyle(
	slide: PptxSlide | undefined,
	slideSize?: SlideBackgroundSize,
): CSSProperties {
	const style = getSlideBackgroundStyle(slide, slideSize);
	// The `--pptx-slide-bg-*` custom properties a background-filled shape paints with.
	const variables = Object.fromEntries(
		Object.entries(style).filter(([key]) => key.startsWith('--')),
	) as CSSProperties;
	return {
		...variables,
		backgroundColor: style['background-color'] as CSSProperties['backgroundColor'],
		backgroundImage: slide?.backgroundImage
			? undefined
			: (style['background-image'] as CSSProperties['backgroundImage']),
		backgroundSize: style['background-size'] as CSSProperties['backgroundSize'],
		backgroundRepeat: style['background-repeat'] as CSSProperties['backgroundRepeat'],
	};
}
