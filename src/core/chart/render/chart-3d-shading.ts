// Office's flat chart-space shading of a 3-D chart's box faces, one darkening amount per face
// orientation. Extracted from `src/ui/src/pptx/render/chart-3d-shading.ts` (PROVENANCE.md). The
// multipliers were measured off Office renders: the front face keeps the base colour, the top is
// about 0.75x and a side about 0.64x.
import { shade } from '../style-palette';

/** A box's six face orientations in the chart's own (unrotated) frame. */
export type Chart3DFace = 'front' | 'back' | 'top' | 'bottom' | 'left' | 'right';

const FACE_SHADE_AMOUNT: Record<Chart3DFace, number> = {
	front: 0,
	top: 0.25,
	right: 0.36,
	left: 0.36,
	back: 0.36,
	bottom: 0.5,
};

const HEX = /^#[0-9a-f]{6}$/i;

/** The base colour darkened for a face. Paints that are not `#RRGGBB` (gradient URLs) are kept. */
export function shadeChart3DFace(baseColor: string, face: Chart3DFace): string {
	const amount = FACE_SHADE_AMOUNT[face];
	return amount === 0 || !HEX.test(baseColor) ? baseColor : shade(baseColor, amount);
}

/**
 * The six face colours in the order three.js `BoxGeometry` expects its material array
 * (`+x, -x, +y, -y, +z, -z`): right, left, top, bottom, front, back.
 */
export function buildChart3DBoxFaceColors(
	baseColor: string,
): readonly [string, string, string, string, string, string] {
	return [
		shadeChart3DFace(baseColor, 'right'),
		shadeChart3DFace(baseColor, 'left'),
		shadeChart3DFace(baseColor, 'top'),
		shadeChart3DFace(baseColor, 'bottom'),
		shadeChart3DFace(baseColor, 'front'),
		shadeChart3DFace(baseColor, 'back'),
	];
}
