import {
	SHAPE_EFFECT_SECTIONS,
	SHAPE_LINE_SECTIONS,
	shapeOutlineChange,
	shapeQuickStylePatch,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxThemeColorRef, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';
import React from 'react';

import { DEFAULT_STROKE_COLOR } from '../../constants';
import { useTranslation } from '../../i18n/use-translation';
import { normalizeHexColor } from '../../utils';
import type { FillStrokePart } from './FillStrokeHelpers';
import { ColorPickerRow } from './FillStrokeSubComponents';
import { InspectorFields } from './InspectorFields';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface StrokeEffectsSectionProps {
	style: ShapeStyle | undefined;
	onUpdateShapeStyle: (updates: Partial<ShapeStyle>) => void;
	onSetStrokeColor: (color: string, ref?: PptxThemeColorRef) => void;
	/**
	 * The Format pane section being drawn: `outline` the line, `effects` the
	 * effects; omitted, both (the fill is `ShapeFillSection`).
	 */
	part?: FillStrokePart;
}

/** The style the effects descriptors read, on any element (for a live colour drag). */
function shapeStyleOf(el: PptxElement): ShapeStyle | undefined {
	return hasShapeProperties(el) ? el.shapeStyle : undefined;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function StrokeEffectsSection({
	style,
	onUpdateShapeStyle,
	onSetStrokeColor,
	part,
}: StrokeEffectsSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const show = (which: FillStrokePart) => part === undefined || part === which;

	return (
		<>
			{/* Stroke color + swatches + recent colors */}
			{show('outline') && (
				<ColorPickerRow
					label={t('pptx.inspector.stroke')}
					prefix='stroke'
					value={normalizeHexColor(style?.strokeColor, DEFAULT_STROKE_COLOR)}
					selectedRef={style?.strokeColorRef}
					onChange={onSetStrokeColor}
					styleChange={(hex, current) => shapeOutlineChange(hex, undefined, current)}
				/>
			)}

			{/* Line options (width, transparency, dash, compound, join, cap) and the
			    Format Shape > Effects sections, drawn from the shared descriptors. */}
			{show('outline') && (
				<InspectorFields
					sections={SHAPE_LINE_SECTIONS}
					target={style}
					onPatch={onUpdateShapeStyle}
				/>
			)}
			{show('effects') && (
				<InspectorFields
					sections={SHAPE_EFFECT_SECTIONS}
					target={style}
					onPatch={onUpdateShapeStyle}
					liveTarget={shapeStyleOf}
					livePatch={shapeQuickStylePatch}
				/>
			)}
		</>
	);
}
