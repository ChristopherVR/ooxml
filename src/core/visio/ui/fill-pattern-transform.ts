import type { VisioMatrix } from '../model';
export {
	composeAffine as composeVisioTransform,
	IDENTITY_AFFINE as VISIO_IDENTITY_TRANSFORM,
} from '../../geometry/affine';

/** Native hatch axes stay parallel to the page, independently of shape/group rotation. */
export function visioFillPatternTransform(
	world: VisioMatrix,
	pageHeight = 0,
): VisioMatrix | undefined {
	const determinant = world[0] * world[3] - world[1] * world[2];
	if (
		!Number.isFinite(determinant) ||
		Math.abs(determinant) < 1e-12 ||
		!Number.isFinite(pageHeight) ||
		pageHeight < 0
	)
		return undefined;
	// Native exported geometry uses a bottom-based page-height origin for its tile phase.
	return [
		world[3] / determinant,
		-world[1] / determinant,
		world[2] / determinant,
		-world[0] / determinant,
		0,
		pageHeight,
	];
}
