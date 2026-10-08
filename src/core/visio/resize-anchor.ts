import type { VisioMatrix } from './model';
import { fail } from './package-common';

/** Fixed normalized local bounds point. Coordinates use the shape's upward local axes. */
export interface VisioResizeAnchor {
	x: 0 | 0.5 | 1;
	y: 0 | 0.5 | 1;
}
export function snapshotResizeAnchor(value: unknown): VisioResizeAnchor | undefined {
	if (value === undefined) return undefined;
	if (!value || typeof value !== 'object') fail('INVALID_EDIT', 'Invalid resize anchor.');
	const { x, y } = value as { x?: unknown; y?: unknown };
	if (![0, 0.5, 1].includes(x as number) || ![0, 0.5, 1].includes(y as number))
		fail('INVALID_EDIT', 'Resize anchor coordinates must be 0, 0.5 or 1.');
	return { x: x as VisioResizeAnchor['x'], y: y as VisioResizeAnchor['y'] };
}
/** Native proportional local-pin geometry, shared by source proof and scene previews.
 * Source admission must prove that retained guarded local pins actually produce these values.
 */
export function visioAnchoredResizeGeometry(
	current: { width: number; height: number; pinX: number; pinY: number; transform: VisioMatrix },
	size: { width: number; height: number },
	anchor: VisioResizeAnchor,
):
	| { pinX: number; pinY: number; locPinX: number; locPinY: number; transform: VisioMatrix }
	| undefined {
	const [a, b, c, d, e, f] = current.transform;
	const determinant = a * d - b * c;
	if (
		![0, 0.5, 1].includes(anchor.x) ||
		![0, 0.5, 1].includes(anchor.y) ||
		![
			current.width,
			current.height,
			current.pinX,
			current.pinY,
			...current.transform,
			size.width,
			size.height,
			anchor.x,
			anchor.y,
		].every(Number.isFinite) ||
		!(current.width > 0 && current.height > 0 && size.width > 0 && size.height > 0) ||
		Math.abs(Math.abs(determinant) - 1) > 1e-8 ||
		Math.abs(a * a + b * b - 1) > 1e-8 ||
		Math.abs(c * c + d * d - 1) > 1e-8 ||
		Math.abs(a * c + b * d) > 1e-8
	)
		return undefined;
	const offsetX = current.pinX - e,
		offsetY = current.pinY - f;
	const oldX = (d * offsetX - c * offsetY) / determinant;
	const oldY = (-b * offsetX + a * offsetY) / determinant;
	const locPinX = (oldX / current.width) * size.width,
		locPinY = (oldY / current.height) * size.height;
	const deltaX = anchor.x * (current.width - size.width) - oldX + locPinX;
	const deltaY = anchor.y * (current.height - size.height) - oldY + locPinY;
	const pinX = current.pinX + a * deltaX + c * deltaY,
		pinY = current.pinY + b * deltaX + d * deltaY;
	const matrix: VisioMatrix = [
		a,
		b,
		c,
		d,
		pinX - a * locPinX - c * locPinY,
		pinY - b * locPinX - d * locPinY,
	];
	return [pinX, pinY, locPinX, locPinY, ...matrix].every(Number.isFinite)
		? { pinX, pinY, locPinX, locPinY, transform: matrix }
		: undefined;
}
