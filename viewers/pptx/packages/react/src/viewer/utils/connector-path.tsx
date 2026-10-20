import { markerPath } from 'ooxml-ui/pptx';
import {
	ConnectorArrowType,
	getConnectorAdjustment as getCoreConnectorAdjustment,
	getConnectorPathGeometry as getCoreConnectorPathGeometry,
	PptxElementWithShapeStyle,
} from 'pptx-viewer-core';
import React from 'react';

import { ConnectorPathGeometry } from '../types';

// Connection-site geometry + compound-line helpers now live in
// `pptx-viewer-shared`; re-export them here to preserve the historical import
// surface for downstream connector code.
export { getConnectionSites, getCompoundLineOffsets, getCompoundLineWidths } from 'ooxml-ui/pptx';

export function getConnectorAdjustment(
	element: PptxElementWithShapeStyle,
	key: string,
	fallback: number,
): number {
	return getCoreConnectorAdjustment(element, key, fallback);
}

export function getConnectorPathGeometry(
	element: PptxElementWithShapeStyle,
): ConnectorPathGeometry {
	return getCoreConnectorPathGeometry(element);
}

export function renderConnectorMarker(
	markerId: string,
	arrowType: ConnectorArrowType | undefined,
	color: string,
	arrowWidth?: 'sm' | 'med' | 'lg',
	arrowLength?: 'sm' | 'med' | 'lg',
): React.ReactNode {
	if (!arrowType || arrowType === 'none') {
		return null;
	}

	const shape = markerPath(arrowType, arrowWidth, arrowLength);

	return (
		<marker
			id={markerId}
			markerWidth={shape.markerWidth}
			markerHeight={shape.markerHeight}
			refX={shape.refX}
			refY={5}
			orient='auto-start-reverse'
			viewBox='0 0 10 10'
			markerUnits='strokeWidth'
			preserveAspectRatio='none'
		>
			{shape.shape === 'circle' ? (
				<circle cx={5} cy={5} r={4} fill={color} />
			) : (
				<path
					d={shape.d}
					fill={shape.strokeOnly ? 'none' : color}
					stroke={shape.strokeOnly ? color : undefined}
				/>
			)}
		</marker>
	);
}
