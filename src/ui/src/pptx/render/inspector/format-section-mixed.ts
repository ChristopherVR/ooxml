/**
 * Mixed values in a multi-selection Format pane: a field (or a whole section)
 * whose value differs across the selected elements shows a "Mixed" tag
 * instead of pretending the primary element's value applies to all.
 *
 * @module render/inspector/format-section-mixed
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties, hasTextProperties } from 'ooxml-core/pptx';

import type { FormatSectionId } from './format-sections';

/** The shared value of a field, or `mixed` when the elements disagree. */
export type FormatFieldValue<V> = { mixed: false; value: V | undefined } | { mixed: true };

/** Compare `read(element)` across `elements` (structural equality for objects). */
export function formatMixedValue<V>(
	elements: readonly PptxElement[],
	read: (element: PptxElement) => V | undefined,
): FormatFieldValue<V> {
	if (elements.length === 0) {
		return { mixed: false, value: undefined };
	}
	const first = read(elements[0]!);
	const key = stable(first);
	for (let i = 1; i < elements.length; i++) {
		if (stable(read(elements[i]!)) !== key) {
			return { mixed: true };
		}
	}
	return { mixed: false, value: first };
}

function stable(value: unknown): string {
	if (value === undefined) {
		return 'u';
	}
	return JSON.stringify(value, (_key, v: unknown) =>
		v && typeof v === 'object' && !Array.isArray(v)
			? Object.fromEntries(Object.entries(v as Record<string, unknown>).sort())
			: v,
	);
}

const shape = (el: PptxElement) => (hasShapeProperties(el) ? el.shapeStyle : undefined);
const text = (el: PptxElement) => (hasTextProperties(el) ? el.textStyle : undefined);

/** What each section compares across a multi-selection. */
const SECTION_READERS: Partial<Record<FormatSectionId, (el: PptxElement) => unknown>> = {
	fill: (el) => {
		const s = shape(el);
		return s && [s.fillMode, s.fillColor, s.fillOpacity, s.fillGradientStops, s.fillPatternPreset];
	},
	outline: (el) => {
		const s = shape(el);
		return s && [s.strokeColor, s.strokeWidth, s.strokeDash, s.strokeOpacity];
	},
	effects: (el) => {
		const s = shape(el);
		return (
			s && [
				s.shadowColor,
				s.shadowBlur,
				s.glowColor,
				s.glowRadius,
				s.softEdgeRadius,
				s.reflectionStartOpacity,
			]
		);
	},
	font: (el) => {
		const t = text(el);
		return t && [t.fontFamily, t.fontSize, t.color, t.bold, t.italic, t.underline];
	},
	paragraph: (el) => {
		const t = text(el);
		return t && [t.align, t.vAlign, t.lineSpacing, t.autoFit];
	},
	arrange: (el) => [el.rotation ?? 0, el.flipHorizontal ?? false, el.flipVertical ?? false],
};

/** Whether a multi-selection disagrees on what `section` edits (always false for one element). */
export function isFormatSectionMixed(
	elements: readonly PptxElement[],
	section: FormatSectionId,
): boolean {
	const read = SECTION_READERS[section];
	return !!read && elements.length > 1 && formatMixedValue(elements, read).mixed;
}

/**
 * The selection the Format pane formats: the primary element first, then the
 * other selected ids found in `elements` (the active slide's, or the edited
 * layout's). Empty when nothing is selected.
 */
export function formatSelectionElements(
	primary: PptxElement | null | undefined,
	ids: readonly string[] | undefined,
	elements: readonly PptxElement[] | undefined,
): PptxElement[] {
	if (!primary) {
		return [];
	}
	const byId = new Map((elements ?? []).map((el) => [el.id, el]));
	const others = (ids ?? [])
		.filter((id) => id !== primary.id)
		.map((id) => byId.get(id))
		.filter((el): el is PptxElement => !!el);
	return [primary, ...others];
}
