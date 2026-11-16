/**
 * The inspector's Alt Text fields (PowerPoint's Alt Text pane): the
 * description (`p:cNvPr/@descr`) and the title (`p:cNvPr/@title`) of the
 * selected element, as field descriptors every binding draws with its one
 * generic renderer.
 *
 * Which element kinds carry which field is `element-non-visual-description.ts`'s
 * decision: a picture models only a description, every other kind both. The
 * descriptors read a `PptxElement` and write a `Partial<PptxElement>`, so a
 * binding sends a change through its usual element patch route. A text field
 * reports what is typed as it is typed (an empty box clears the text), the
 * way PowerPoint's pane applies alt text while you write.
 *
 * @module render/inspector/element-alt-text-fields
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { supportsAltTextField, supportsTitleField } from '../element-non-visual-description';
import type { InspectorSectionDescriptor } from './inspector-field-model';

export type ElementAltTextSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;

function textOf(element: PptxElement, key: 'altText' | 'title'): string {
	const value = (element as Partial<Record<'altText' | 'title', unknown>>)[key];
	return typeof value === 'string' ? value : '';
}

/** The description and title fields (untitled: each pane heads them). */
export const ELEMENT_ALT_TEXT_SECTION: ElementAltTextSection = {
	id: 'altText',
	fields: [
		{
			id: 'altText',
			kind: 'text',
			multiline: true,
			labelKey: 'pptx.elementAccessibility.altText',
			placeholderKey: 'pptx.elementAccessibility.altTextPlaceholder',
			wide: true,
			testId: 'element-alt-text',
			visibleWhen: (element) => supportsAltTextField(element.type),
			read: (element) => textOf(element, 'altText'),
			write: (_element, value) => ({ altText: String(value) }) as Partial<PptxElement>,
		},
		{
			id: 'title',
			kind: 'text',
			labelKey: 'pptx.elementAccessibility.title',
			placeholderKey: 'pptx.elementAccessibility.titlePlaceholder',
			wide: true,
			testId: 'element-alt-title',
			visibleWhen: (element) => supportsTitleField(element.type),
			read: (element) => textOf(element, 'title'),
			write: (_element, value) => ({ title: String(value) }) as Partial<PptxElement>,
		},
	],
};

/** The Alt Text fields as a section list, what a generic renderer takes. */
export const ELEMENT_ALT_TEXT_SECTIONS: readonly ElementAltTextSection[] = [
	ELEMENT_ALT_TEXT_SECTION,
];
