import type { VisioFillGradient, VisioGradientRegion } from './model';
import { linearGradientEndpoints } from './theme-gradient';
import { radialFillGradient } from './radial-fill-gradient';
import { number, type Cells } from './sheet';
import { clampUnitInterval } from '../color/color-primitives';

/** Native cached classic gradients 25-40. */
export function legacyFillGradient(
	cells: Cells,
	width: number,
	height: number,
	resolve: (name: string) => string,
): VisioFillGradient | undefined {
	const pattern = number(cells, 'FillPattern', 1);
	if (
		pattern < 25 ||
		pattern > 40 ||
		!Number.isInteger(pattern) ||
		!Number.isFinite(width) ||
		width <= 0 ||
		!Number.isFinite(height) ||
		height <= 0 ||
		number(cells, 'FillGradientEnabled', 0) !== 0
	)
		return undefined;
	const foreground = resolve('FillForegnd'),
		background = resolve('FillBkgnd');
	if (!foreground || !background) return undefined;
	const front = {
		color: foreground,
		opacity: clampUnitInterval(1 - number(cells, 'FillForegndTrans', 0)),
	};
	const back = {
		color: background,
		opacity: clampUnitInterval(1 - number(cells, 'FillBkgndTrans', 0)),
	};
	if (pattern >= 31 && pattern <= 35) {
		// Native SVG uses triangle-local bounding-box linear gradients.
		const triangle = (
			points: VisioGradientRegion['points'],
			angle: number,
		): VisioGradientRegion => ({
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
			stops: [
				{ offset: 0, ...front },
				{ offset: 1, ...back },
			],
		};
	}
	if (pattern >= 36)
		return radialFillGradient([7, 6, 2, 1, 3][pattern - 36]!, [
			{ offset: 0, ...front },
			{ offset: 1, ...back },
		]);
	const angle = pattern <= 26 ? 0 : pattern === 27 ? 180 : pattern <= 29 ? 90 : 270;
	return {
		type: 'linear',
		...linearGradientEndpoints(width, height, angle * 60_000),
		stops:
			pattern === 26 || pattern === 29
				? [
						{ offset: 0, ...back },
						{ offset: 0.5, ...front },
						{ offset: 1, ...back },
					]
				: [
						{ offset: 0, ...front },
						{ offset: 1, ...back },
					],
	};
}
