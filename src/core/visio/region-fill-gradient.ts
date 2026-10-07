import type { VisioGradientRegion, VisioRegionGradient } from './model';
import { linearGradientEndpoints } from './theme-gradient';

/** Native rectangular presets, shared by saved gradients and classic fill patterns. */
export function regionFillGradient(
	direction: number,
	stops: VisioRegionGradient['stops'],
): VisioRegionGradient | undefined {
	const pattern = Number.isInteger(direction) ? [34, 33, 35, 32, 31][direction - 8] : undefined;
	if (pattern === undefined) return undefined;
	// Native SVG uses triangle-local bounding-box linear gradients.
	const triangle = (points: VisioGradientRegion['points'], angle: number): VisioGradientRegion => ({
		points,
		angle: angle === 0 ? 360 : angle,
		...linearGradientEndpoints(1, 1, angle * 60_000),
	});
	const rising: VisioGradientRegion['points'][] = [
		[
			[0, 1],
			[0, 0],
			[1, 0],
		],
		[
			[0, 1],
			[1, 1],
			[1, 0],
		],
	];
	const falling: VisioGradientRegion['points'][] = [
		[
			[0, 0],
			[0, 1],
			[1, 1],
		],
		[
			[0, 0],
			[1, 0],
			[1, 1],
		],
	];
	const regions =
		pattern === 35
			? [
					triangle(
						[
							[0.5, 0.5],
							[0, 1],
							[0, 0],
						],
						180,
					),
					triangle(
						[
							[0.5, 0.5],
							[1, 1],
							[1, 0],
						],
						0,
					),
					triangle(
						[
							[0.5, 0.5],
							[0, 1],
							[1, 1],
						],
						270,
					),
					triangle(
						[
							[0.5, 0.5],
							[0, 0],
							[1, 0],
						],
						90,
					),
				]
			: pattern === 31
				? [triangle(rising[0]!, 90), triangle(rising[1]!, 0)]
				: pattern === 32
					? [triangle(falling[0]!, 180), triangle(falling[1]!, 90)]
					: pattern === 33
						? [triangle(falling[0]!, 270), triangle(falling[1]!, 0)]
						: [triangle(rising[1]!, 270), triangle(rising[0]!, 180)];
	return {
		type: 'regions',
		regions,
		stops,
	};
}
