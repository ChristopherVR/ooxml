/**
 * The element Size & Position fields of the inspector: X, Y, width, height
 * and rotation, as one untitled section of field descriptors every binding
 * draws with its one generic renderer (the lock toggle beside them stays per
 * binding).
 *
 * The descriptors read the element's frame (in the editor's px units) and
 * produce a geometry patch, which a binding routes through its own geometry
 * commit (an element patch, or the vanilla binding's `setGeometry`).
 *
 * @module render/inspector/element-geometry-fields
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { InspectorFieldDescriptor, InspectorSectionDescriptor } from './inspector-field-model';

/** The frame the Size & Position fields read. */
export type ElementGeometryTarget = Pick<PptxElement, 'x' | 'y' | 'width' | 'height'> & {
	rotation?: number | undefined;
};

/** The patch a Size & Position field produces. */
export type ElementGeometryPatch = Partial<
	Pick<PptxElement, 'x' | 'y' | 'width' | 'height' | 'rotation'>
>;

export type ElementGeometrySection = InspectorSectionDescriptor<
	ElementGeometryTarget,
	ElementGeometryPatch
>;
type GeometryField = InspectorFieldDescriptor<ElementGeometryTarget, ElementGeometryPatch>;

/** Smallest width or height the fields accept, in px. */
export const MIN_ELEMENT_SIZE = 1;

function geometryField(
	id: 'x' | 'y' | 'width' | 'height' | 'rotation',
	labelKey: string,
	min?: number,
): GeometryField {
	const field: GeometryField = {
		id,
		kind: 'number',
		labelKey,
		decimals: 0,
		testId: `geometry-${id}`,
		read: (el) => el[id] ?? 0,
		write: (_el, v) => ({ [id]: Number(v) }),
	};
	if (min !== undefined) field.min = min;
	return field;
}

/** The Size & Position section (no title: each pane heads it with its lock toggle). */
export const ELEMENT_GEOMETRY_SECTION: ElementGeometrySection = {
	id: 'geometry',
	fields: [
		geometryField('x', 'pptx.inspector.x'),
		geometryField('y', 'pptx.inspector.y'),
		geometryField('width', 'pptx.inspector.w', MIN_ELEMENT_SIZE),
		geometryField('height', 'pptx.inspector.h', MIN_ELEMENT_SIZE),
		geometryField('rotation', 'pptx.inspector.rotation'),
	],
};

/** The Size & Position fields as a section list, what a generic renderer takes. */
export const ELEMENT_GEOMETRY_SECTIONS: readonly ElementGeometrySection[] = [
	ELEMENT_GEOMETRY_SECTION,
];
