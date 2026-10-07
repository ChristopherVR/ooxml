import type { VisioLinearGradient } from './model';
import { linearGradientEndpoints } from './theme-gradient';
import { number, type Cells } from './sheet';
import { clampUnitInterval } from '../color/color-primitives';

/** Native cached orthogonal patterns 25-30, including symmetric center-color patterns. */
export function legacyFillGradient(
	cells: Cells,
	width: number,
	height: number,
	resolve: (name: string) => string,
): VisioLinearGradient | undefined {
	const pattern = number(cells, 'FillPattern', 1);
	if (
		pattern < 25 ||
		pattern > 30 ||
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
