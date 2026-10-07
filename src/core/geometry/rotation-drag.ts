export interface RotationPoint {
	x: number;
	y: number;
}
/** Clockwise screen degrees, with zero pointing up and a result in [0, 360). */
export function computeRotation(center: RotationPoint, pointer: RotationPoint): number {
	const deg = Math.atan2(pointer.x - center.x, center.y - pointer.y) * (180 / Math.PI);
	return ((deg % 360) + 360) % 360;
}
/**
 * Anchor a rotate drag to the actual press, including its offset inside the
 * handle's hit area. Points must share one coordinate space. Snapping and
 * rounding remain the caller's existing policy.
 */
export function createRotationDrag(
	center: RotationPoint,
	startPointer: RotationPoint,
	startRotation: number,
	options: { unwrapped?: boolean } = {},
): (pointer: RotationPoint) => number {
	const bearing = (pointer: RotationPoint): number | undefined =>
		pointer.x === center.x && pointer.y === center.y ? undefined : computeRotation(center, pointer);
	let startAngle = bearing(startPointer);
	let rotation = startRotation;
	let previousAngle = startAngle;
	return (pointer) => {
		const angle = bearing(pointer);
		// A zero-height connector can place its rotate handle at its center.
		// There is no bearing there: wait for a direction instead of jumping.
		if (angle === undefined) {
			return rotation;
		}
		if (startAngle === undefined) {
			startAngle = angle;
			previousAngle = angle;
			return rotation;
		}
		if (options.unwrapped) {
			const delta = ((angle - (previousAngle ?? angle) + 540) % 360) - 180;
			rotation += delta;
			previousAngle = angle;
		} else rotation = (((startRotation + angle - startAngle) % 360) + 360) % 360;
		return rotation;
	};
}
