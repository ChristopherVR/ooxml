import React from 'react';

import { isLineish } from './FillStrokeHelpers';
import type { FillStrokePropertiesProps } from './FillStrokeHelpers';
import { QuickStylesGallery } from './QuickStylesGallery';
import { ShapeFillSection } from './ShapeFillSection';
import { StrokeEffectsSection } from './StrokeEffectsSection';

export type { FillStrokePropertiesProps };

/**
 * Inspector panel for editing the fill and stroke properties of shapes and text elements.
 *
 * The Fill part is the shared fill descriptors and actions ({@link ShapeFillSection}),
 * after the quick style presets for shapes and text; the Outline and Effects parts
 * are {@link StrokeEffectsSection}.
 *
 * @param props - {@link FillStrokePropertiesProps}
 * @returns The fill/stroke inspector panel.
 */
export function FillStrokeProperties({
	selectedElement,
	selectedShapeStyle,
	selectedShapeType,
	onUpdateShapeStyle,
	onSetFillColor,
	onSetStrokeColor,
	part,
}: FillStrokePropertiesProps): React.ReactElement {
	// Lines (and line-like shapes) have fill disabled since they have no area to fill
	const line = isLineish(selectedElement, selectedShapeType);
	const style = selectedShapeStyle;
	const showQuickStyles = selectedElement.type === 'shape' || selectedElement.type === 'text';
	const fill = part === undefined || part === 'fill';
	const strokeParts = part !== 'fill' && (
		<StrokeEffectsSection
			style={style}
			onUpdateShapeStyle={onUpdateShapeStyle}
			onSetStrokeColor={onSetStrokeColor}
			part={part}
		/>
	);
	if (!fill) {
		return <div className='grid grid-cols-2 gap-2'>{strokeParts}</div>;
	}

	return (
		<>
			{showQuickStyles && <QuickStylesGallery onUpdateShapeStyle={onUpdateShapeStyle} />}

			<div className='grid grid-cols-2 gap-2'>
				<ShapeFillSection
					style={style}
					isLine={line}
					onUpdateShapeStyle={onUpdateShapeStyle}
					onSetFillColor={onSetFillColor}
				/>
				{strokeParts}
			</div>
		</>
	);
}
