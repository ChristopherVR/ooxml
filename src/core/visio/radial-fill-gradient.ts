import type { VisioRadialGradient } from './model';

/** Native SVG presets in normalized local y-up coordinates, shared by saved and classic fills. */
export function radialFillGradient(
	direction: number,
	stops: VisioRadialGradient['stops'],
): VisioRadialGradient | undefined {
	const presets: readonly (readonly [number, number, number])[] = [
		[1, 0, 1.4],
		[0, 0, 1.4],
		[0.5, 0.5, 0.73],
		[0.5, 1, 1.1],
		[0.5, 0, 1.1],
		[1, 1, 1.4],
		[0, 1, 1.4],
	];
	const preset = Number.isInteger(direction) ? presets[direction - 1] : undefined;
	return preset
		? { type: 'radial', center: [preset[0], preset[1]], radius: preset[2], stops }
		: undefined;
}
