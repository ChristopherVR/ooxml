/** PowerPoint adapter for the shared, strict chart marker geometry. */
import {
	buildMarkerPrimitive as sharedMarkerPrimitive,
	type MarkerShapeInput as SharedMarkerShapeInput,
} from 'ooxml-core/chart';
import type { ChartPartRef, SvgCircle, SvgPath, SvgPolygon, SvgRect } from './chart-view-model';

export { automaticMarkerSymbol, markerOutline } from 'ooxml-core/chart';
export type MarkerPrimitive = SvgCircle | SvgRect | SvgPath | SvgPolygon;
export type MarkerShapeInput = SharedMarkerShapeInput<ChartPartRef>;

export function buildMarkerPrimitive(input: MarkerShapeInput): MarkerPrimitive | null {
	return sharedMarkerPrimitive(input);
}
