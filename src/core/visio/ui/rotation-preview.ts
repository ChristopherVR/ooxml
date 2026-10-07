import type { VisioShape, VisioMatrix } from '../model';
import { transform } from '../geometry';
import { composeAffine } from '../../geometry/affine';
/** Rotate a cached pose around its saved parent-space pin without changing the source scene. */
export function visioRotationPreviewTransform(
	shape: VisioShape,
	angle: number,
): VisioMatrix | undefined {
	const pin = shape.rotation;
	if (!pin || ![angle, pin.pinX, pin.pinY, pin.angle, ...shape.transform].every(Number.isFinite))
		return undefined;
	const delta = angle - pin.angle;
	if (delta === 0) return shape.transform;
	return composeAffine(transform(pin.pinX, pin.pinY, pin.pinX, pin.pinY, delta), shape.transform);
}
