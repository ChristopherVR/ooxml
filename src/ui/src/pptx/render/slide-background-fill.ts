/**
 * A shape painted with the slide background (`<p:sp useBgFill="1">`, the Fill
 * pane's "Slide background fill"), drawn live.
 *
 * The load pipeline resolves such a shape against its slide's background once
 * (`applyBackgroundFillToShapes` in the core), but a shape switched to it in
 * the inspector has no fill of its own until the deck is reopened, and the
 * fill resolver sees only the element. So the slide stage publishes its
 * resolved background as CSS custom properties ({@link slideBackgroundVariables},
 * merged into `getSlideBackgroundStyle`), and the fill resolver paints a
 * background-filled shape with them ({@link slideBackgroundShapeFill}). The
 * shape inherits the properties from the stage it is drawn in, so every
 * binding draws it without wiring of its own; the paint is offset by the
 * shape's position so the slide background lines up across the shape, as it
 * does in PowerPoint.
 *
 * @module render/slide-background-fill
 */
import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';

import type { CssStyleMap } from './element-style-transform';

/** The custom properties a slide stage publishes for background-filled shapes. */
export const SLIDE_BACKGROUND_VARIABLES = {
	color: '--pptx-slide-bg-color',
	image: '--pptx-slide-bg-image',
	size: '--pptx-slide-bg-size',
	repeat: '--pptx-slide-bg-repeat',
} as const;

/**
 * The custom properties for a stage whose resolved `background-*` style is
 * `style`: the same paint, sized to the whole slide when its size is known.
 */
export function slideBackgroundVariables(
	style: CssStyleMap,
	slideSize?: { widthPx: number; heightPx: number },
): CssStyleMap {
	const image = style['background-image'];
	const pattern = style['background-repeat'] === 'repeat';
	const vars: CssStyleMap = {
		[SLIDE_BACKGROUND_VARIABLES.color]: style['background-color'] ?? 'transparent',
		[SLIDE_BACKGROUND_VARIABLES.image]: image ?? 'none',
		[SLIDE_BACKGROUND_VARIABLES.repeat]: pattern ? 'repeat' : 'no-repeat',
	};
	if (!pattern && slideSize && slideSize.widthPx > 0 && slideSize.heightPx > 0) {
		vars[SLIDE_BACKGROUND_VARIABLES.size] = `${slideSize.widthPx}px ${slideSize.heightPx}px`;
	} else if (pattern) {
		vars[SLIDE_BACKGROUND_VARIABLES.size] = 'auto';
	}
	return vars;
}

/** Whether `style` paints the slide background but carries no resolved copy of it. */
export function paintsLiveSlideBackground(style: ShapeStyle | undefined): boolean {
	return style?.useBackgroundFill === true && (style.fillMode ?? 'none') === 'none';
}

/**
 * The CSS fill of a background-filled shape: the stage's published slide
 * background, offset by the shape's position on the slide.
 */
export function slideBackgroundShapeFill(element: PptxElement): {
	backgroundColor: string;
	backgroundImage: string;
	backgroundSize: string;
	backgroundRepeat: string;
	backgroundPosition: string;
} {
	const v = SLIDE_BACKGROUND_VARIABLES;
	return {
		backgroundColor: `var(${v.color}, transparent)`,
		backgroundImage: `var(${v.image}, none)`,
		backgroundSize: `var(${v.size}, 100% 100%)`,
		backgroundRepeat: `var(${v.repeat}, no-repeat)`,
		backgroundPosition: `${-(element.x ?? 0)}px ${-(element.y ?? 0)}px`,
	};
}
