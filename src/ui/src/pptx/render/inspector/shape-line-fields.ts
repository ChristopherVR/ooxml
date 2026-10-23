/**
 * The shape Line options of the inspector (Format Shape > Fill & Line > Line):
 * width, transparency, dash type, compound type, join and cap, as one untitled
 * section of field descriptors every binding draws with its one generic
 * renderer. The line colour stays with each binding's theme colour picker.
 *
 * Like the Effects descriptors these edit a `ShapeStyle` and produce a
 * `Partial<ShapeStyle>`; the option catalogues are the shared
 * `stroke-dash-options.ts` and `stroke-line-style-options.ts` lists.
 *
 * @module render/inspector/shape-line-fields
 */
import type { ShapeStyle } from 'ooxml-core/pptx';

import { STROKE_DASH_OPTIONS } from '../stroke-dash-options';
import {
	COMPOUND_LINE_OPTIONS,
	LINE_CAP_OPTIONS,
	LINE_JOIN_OPTIONS,
} from '../stroke-line-style-options';
import type {
	InspectorFieldDescriptor,
	InspectorFieldOption,
	InspectorFieldValue,
	InspectorSectionDescriptor,
} from './inspector-field-model';

/** A Line field: reads the shape style, writes a shape-style patch. */
type LineField = InspectorFieldDescriptor<ShapeStyle | undefined, Partial<ShapeStyle>>;
export type ShapeLineSection = InspectorSectionDescriptor<
	ShapeStyle | undefined,
	Partial<ShapeStyle>
>;

/** Widest line the width box offers, in px. */
const MAX_LINE_WIDTH = 100;

const toOptions = (
	list: readonly { value: string; i18nKey: string }[],
): readonly InspectorFieldOption[] => list.map((o) => ({ value: o.value, labelKey: o.i18nKey }));

/** A `select` field over one `ShapeStyle` key with its shared catalogue. */
function choice<K extends 'strokeDash' | 'compoundLine' | 'lineJoin' | 'lineCap'>(
	key: K,
	labelKey: string,
	list: readonly { value: string; i18nKey: string }[],
	fallback: NonNullable<ShapeStyle[K]>,
	testId: string,
): LineField {
	return {
		id: key,
		kind: 'select',
		labelKey,
		options: toOptions(list),
		testId,
		read: (s) => s?.[key] ?? fallback,
		write: (_s, v: InspectorFieldValue) => ({ [key]: String(v) }) as Partial<ShapeStyle>,
	};
}

/** The Line options section (no title: each pane heads it with its line colour). */
export const SHAPE_LINE_SECTION: ShapeLineSection = {
	id: 'line',
	fields: [
		{
			id: 'strokeWidth',
			kind: 'length',
			labelKey: 'pptx.strokeEffects.strokeWidth',
			min: 0,
			max: MAX_LINE_WIDTH,
			step: 0.5,
			testId: 'line-width',
			read: (s) => s?.strokeWidth ?? 0,
			write: (_s, v) => ({ strokeWidth: Number(v) }),
		},
		{
			id: 'strokeOpacity',
			kind: 'percent',
			labelKey: 'pptx.strokeEffects.strokeOpacity',
			unit: '%',
			testId: 'line-opacity',
			read: (s) => (s?.strokeOpacity ?? 1) * 100,
			write: (_s, v) => ({ strokeOpacity: Number(v) / 100 }),
		},
		choice(
			'strokeDash',
			'pptx.strokeEffects.strokeDash',
			STROKE_DASH_OPTIONS,
			'solid',
			'line-dash',
		),
		choice(
			'compoundLine',
			'pptx.strokeEffects.compoundLine',
			COMPOUND_LINE_OPTIONS,
			'sng',
			'line-compound',
		),
		choice('lineJoin', 'pptx.strokeEffects.lineJoin', LINE_JOIN_OPTIONS, 'round', 'line-join'),
		choice('lineCap', 'pptx.strokeEffects.lineCap', LINE_CAP_OPTIONS, 'flat', 'line-cap'),
	],
};

/** The Line options as a section list, what a generic renderer takes. */
export const SHAPE_LINE_SECTIONS: readonly ShapeLineSection[] = [SHAPE_LINE_SECTION];
