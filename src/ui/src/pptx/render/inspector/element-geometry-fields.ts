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
import type { PptxElement, SummaryZoomTarget, ZoomPptxElement } from 'ooxml-core/pptx';

import type { InspectorFieldDescriptor, InspectorSectionDescriptor } from './inspector-field-model';

/** The frame the Size & Position fields read. */
export type ElementGeometryTarget = Pick<PptxElement, 'x' | 'y' | 'width' | 'height'> & {
	rotation?: number | undefined;
	type?: PptxElement['type'] | undefined;
	zoomType?: ZoomPptxElement['zoomType'] | undefined;
	returnToParent?: boolean | undefined;
	zoomShowBackground?: boolean | undefined;
	transitionDurationMs?: number | undefined;
	summaryTargets?: SummaryZoomTarget[] | undefined;
};

/** The patch a Size & Position field produces. */
export type ElementGeometryPatch = Partial<
	Pick<
		ElementGeometryTarget,
		| 'x'
		| 'y'
		| 'width'
		| 'height'
		| 'rotation'
		| 'returnToParent'
		| 'zoomShowBackground'
		| 'transitionDurationMs'
		| 'summaryTargets'
	> &
		Pick<ZoomPptxElement, 'zoomImageType' | 'imagePath' | 'imageData'>
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

type ZoomOptions = Pick<
	ZoomPptxElement,
	'returnToParent' | 'zoomShowBackground' | 'transitionDurationMs'
>;
const ZOOM_OPTION_FIELDS: readonly InspectorFieldDescriptor<ZoomOptions, ZoomOptions>[] = [
	{
		id: 'zoomReturn',
		kind: 'toggle',
		labelKey: 'pptx.graphic.zoomReturn',
		read: (zoom) => zoom.returnToParent !== false,
		write: (_zoom, value) => ({ returnToParent: value === true }),
	},
	{
		id: 'zoomTransparent',
		kind: 'toggle',
		labelKey: 'pptx.graphic.zoomTransparent',
		read: (zoom) => zoom.zoomShowBackground === false,
		write: (_zoom, value) => ({ zoomShowBackground: value !== true }),
	},
	{
		id: 'zoomDuration',
		kind: 'number',
		labelKey: 'pptx.transition.duration',
		min: 0,
		max: 60000,
		decimals: 0,
		read: (zoom) => zoom.transitionDurationMs ?? 0,
		write: (_zoom, value) => ({ transitionDurationMs: Number(value) }),
	},
];
export const ZOOM_OPTIONS_SECTION: ElementGeometrySection = {
	id: 'zoomOptions',
	labelKey: 'pptx.graphic.zoomOptions',
	fields: ZOOM_OPTION_FIELDS,
};
const ZOOM_GEOMETRY_SECTIONS: readonly ElementGeometrySection[] = [
	ELEMENT_GEOMETRY_SECTION,
	ZOOM_OPTIONS_SECTION,
];

function summaryTileSections(targets: readonly SummaryZoomTarget[]): ElementGeometrySection[] {
	return targets.map((target, index): ElementGeometrySection => ({
		id: `zoomTile-${index}`,
		label: target.title?.trim() || `#${target.targetSlideIndex + 1}`,
		fields: ZOOM_OPTION_FIELDS.map((field): GeometryField => ({
			...field,
			read: (zoom) => field.read(zoom.summaryTargets?.[index] ?? {}),
			write: (zoom, value) => ({
				summaryTargets: zoom.summaryTargets?.map((tile, tileIndex) =>
					tileIndex === index ? { ...tile, ...field.write(tile, value) } : tile,
				),
			}),
		})),
	}));
}

/** Playback and appearance settings for a Slide Zoom, Section Zoom or individual Summary tile. */
export function sizePositionSections(element: ElementGeometryTarget | undefined) {
	if (element?.type !== 'zoom') return ELEMENT_GEOMETRY_SECTIONS;
	return element.zoomType === 'summary'
		? [...ELEMENT_GEOMETRY_SECTIONS, ...summaryTileSections(element.summaryTargets ?? [])]
		: ZOOM_GEOMETRY_SECTIONS;
}
