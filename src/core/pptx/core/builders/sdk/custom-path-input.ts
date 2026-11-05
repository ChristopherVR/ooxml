/**
 * Custom geometry of a new shape (`a:custGeom`) from SVG path data or
 * segments, one `a:path` per input path.
 *
 * @module sdk/custom-path-input
 */

import { svgPathToSegments } from '../../../../geometry/svg-path-segments';
import { customGeometryPathsToSvg } from '../../geometry/custom-geometry';
import type { ShapePptxElement } from '../../types/elements';
import type {
	CustomGeometryPath,
	CustomGeometryPoint,
	CustomGeometrySegment,
} from '../../types/geometry';
import type { CustomPathInput } from './types';

const EMU_PER_PX = 9525;

type CustomGeometryFields = Pick<
	ShapePptxElement,
	'shapeType' | 'pathData' | 'pathWidth' | 'pathHeight' | 'customGeometryPaths'
>;

/**
 * The custom geometry fields of a shape `width` x `height` px from `paths`.
 * Each path's `width` x `height` box is stretched over the shape, as
 * DrawingML does with `a:path/@w`, `@h`; the paths are stored in EMU of the
 * shape's size so their coordinates keep their precision when the writer
 * rounds them to integers.
 */
export function customGeometryFromInput(
	paths: readonly CustomPathInput[],
	width: number,
	height: number,
): CustomGeometryFields {
	const w = Math.max(1, Math.round(width * EMU_PER_PX));
	const h = Math.max(1, Math.round(height * EMU_PER_PX));
	const customGeometryPaths = paths.map((input): CustomGeometryPath => {
		if (!(input.width > 0) || !(input.height > 0) || !Number.isFinite(input.width * input.height)) {
			throw new RangeError('CustomPathInput width and height must be positive and finite');
		}
		if (input.d !== undefined && input.segments !== undefined) {
			throw new RangeError('CustomPathInput takes d or segments, not both');
		}
		const segments = input.segments ?? (input.d !== undefined ? svgPathToSegments(input.d) : []);
		if (segments.length === 0) {
			throw new RangeError('CustomPathInput needs a non-empty d or segments');
		}
		if (!segments.every(isFiniteSegment)) {
			throw new RangeError('CustomPathInput segments must have finite coordinates');
		}
		const sx = w / input.width;
		const sy = h / input.height;
		const path: CustomGeometryPath = {
			width: w,
			height: h,
			segments: segments.map((segment) => scaleSegment(segment, sx, sy)),
		};
		if (input.fill !== undefined) {
			path.fillMode = input.fill;
		}
		if (input.stroke !== undefined) {
			path.stroke = input.stroke;
		}
		return path;
	});
	return {
		shapeType: 'custom',
		pathData: customGeometryPathsToSvg(customGeometryPaths),
		pathWidth: w,
		pathHeight: h,
		customGeometryPaths,
	};
}

function isFiniteSegment(segment: CustomGeometrySegment): boolean {
	switch (segment.type) {
		case 'moveTo':
		case 'lineTo':
			return Number.isFinite(segment.pt.x) && Number.isFinite(segment.pt.y);
		case 'cubicBezTo':
		case 'quadBezTo':
			return segment.pts.every((pt) => Number.isFinite(pt.x) && Number.isFinite(pt.y));
		case 'arcTo':
			return [segment.wR, segment.hR, segment.stAng, segment.swAng].every(Number.isFinite);
		case 'close':
			return true;
	}
}

function scaleSegment(
	segment: CustomGeometrySegment,
	sx: number,
	sy: number,
): CustomGeometrySegment {
	const p = (pt: CustomGeometryPoint): CustomGeometryPoint => ({ x: pt.x * sx, y: pt.y * sy });
	switch (segment.type) {
		case 'moveTo':
		case 'lineTo':
			return { type: segment.type, pt: p(segment.pt) };
		case 'cubicBezTo':
			return { type: 'cubicBezTo', pts: [p(segment.pts[0]), p(segment.pts[1]), p(segment.pts[2])] };
		case 'quadBezTo':
			return { type: 'quadBezTo', pts: [p(segment.pts[0]), p(segment.pts[1])] };
		case 'arcTo':
			return { ...segment, wR: segment.wR * sx, hR: segment.hR * sy };
		case 'close':
			return segment;
	}
}
