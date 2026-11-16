/**
 * The inspector's slide Background card fields (PowerPoint's Format
 * Background > Fill: the solid colour and Hide background graphics), as field
 * descriptors every binding draws with its one generic renderer.
 *
 * The descriptors read the active `PptxSlide` and write a `Partial<PptxSlide>`
 * the binding applies through its usual slide route. A solid colour replaces
 * the slide's picture, gradient and pattern background, as picking a solid
 * fill in PowerPoint does. The picture chooser and the Clear Background action
 * (`slideBackgroundClearState`) stay native templates beside them.
 *
 * @module render/inspector/slide-background-fields
 */
import type { PptxSlide } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import type { InspectorSectionDescriptor } from './inspector-field-model';

export type SlideBackgroundTarget = PptxSlide | undefined;
export type SlideBackgroundSection = InspectorSectionDescriptor<
	SlideBackgroundTarget,
	Partial<PptxSlide>
>;

/** The colour a slide with no background of its own shows in the picker. */
const DEFAULT_SLIDE_COLOUR = '#ffffff';

/** The solid colour and Hide background graphics (untitled: the card heads them). */
export const SLIDE_BACKGROUND_SECTION: SlideBackgroundSection = {
	id: 'slideBackground',
	fields: [
		{
			id: 'backgroundColor',
			kind: 'color',
			labelKey: 'pptx.slideBackground.colour',
			wide: true,
			testId: 'slide-background-colour',
			read: (slide) => normalizeHexColor(slide?.backgroundColor, DEFAULT_SLIDE_COLOUR),
			write: (_slide, value) => ({
				backgroundColor: String(value),
				backgroundImage: undefined,
				backgroundGradient: undefined,
				backgroundPattern: undefined,
			}),
		},
		{
			id: 'hideBackgroundGraphics',
			kind: 'toggle',
			labelKey: 'pptx.slideBackground.hideBackgroundGraphics',
			testId: 'slide-hide-background-graphics',
			read: (slide) => slide?.showMasterShapes === false,
			write: (_slide, value) => ({ showMasterShapes: value !== true }),
		},
	],
};

/** The card's fields as a section list, what a generic renderer takes. */
export const SLIDE_BACKGROUND_SECTIONS: readonly SlideBackgroundSection[] = [
	SLIDE_BACKGROUND_SECTION,
];
