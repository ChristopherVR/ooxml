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
import { RNG, DIS, LBL } from './FillStrokeHelpers';
import { ColorPickerRow } from './FillStrokeSubComponents';
import { InspectorFields } from './InspectorFields';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface StrokeEffectsSectionProps {
	style: ShapeStyle | undefined;
	isLine: boolean;
	onUpdateShapeStyle: (updates: Partial<ShapeStyle>) => void;
	onSetStrokeColor: (color: string, ref?: PptxThemeColorRef) => void;
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
	isLine,
	onUpdateShapeStyle,
	onSetStrokeColor,
}: StrokeEffectsSectionProps): React.ReactElement {
	const { t } = useTranslation();

	return (
		<>
			{/* Stroke color + swatches + recent colors */}
			<ColorPickerRow
				label={t('pptx.inspector.stroke')}
				prefix='stroke'
				value={normalizeHexColor(style?.strokeColor, DEFAULT_STROKE_COLOR)}
				selectedRef={style?.strokeColorRef}
				onChange={onSetStrokeColor}
				styleChange={(hex, current) => shapeOutlineChange(hex, undefined, current)}
			/>

			{/* Fill Opacity */}
			<label className='flex flex-col gap-1'>
				<span className={LBL}>{t('pptx.strokeEffects.fillOpacity')}</span>
				<input
					type='range'
					min={0}
					max={100}
					disabled={isLine}
					value={Math.round((style?.fillOpacity ?? 1) * 100)}
					onChange={(e) => onUpdateShapeStyle({ fillOpacity: Number(e.target.value) / 100 })}
					className={`${RNG} ${DIS}`}
				/>
			</label>

			{/* Line options (width, transparency, dash, compound, join, cap) and the
			    Format Shape > Effects sections, drawn from the shared descriptors. */}
			<InspectorFields sections={SHAPE_LINE_SECTIONS} target={style} onPatch={onUpdateShapeStyle} />
			<InspectorFields
				sections={SHAPE_EFFECT_SECTIONS}
				target={style}
				onPatch={onUpdateShapeStyle}
				liveTarget={shapeStyleOf}
				livePatch={shapeQuickStylePatch}
			/>
		</>
	);
}
