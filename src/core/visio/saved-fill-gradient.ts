import type { VisioFillGradient, VisioLinearGradient, VisioGeometry } from './model';
import { pathFillGradient } from './path-fill-gradient';
import { radialFillGradient } from './radial-fill-gradient';
import { regionFillGradient } from './region-fill-gradient';
import { number, sectionRows, type Cells, type Report, type Sheet } from './sheet';
import { linearGradientEndpoints } from './theme-gradient';
import { canUseVisioSigmaInterpolation } from './native-gradient-stops';

/**
 * Saved ShapeSheet gradients use radians and normalized [0,1] stop values.
 * Only complete local, shape-rotating caches with supported geometry are accepted. Theme and
 * root-style substitution happen before this function; missing caches are not
 * inferred from formulas, legacy pattern numbers, or an unrelated theme.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/fill-gradient-section
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/e19c498a-5277-4add-9953-8b85cb2af250
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/64030657-ec1c-4bcd-adbb-701bb24004fa
 */
export function savedShapeGradient(
	sheet: Sheet,
	width: number,
	height: number,
	resolveColor: (cells: Cells) => string,
	report: Report,
	geometry: readonly VisioGeometry[] = [],
	paint: 'Fill' | 'Line' = 'Fill',
): VisioFillGradient | undefined {
	const cells = sheet.cells;
	if (number(cells, `${paint}GradientEnabled`, NaN) !== 1) return undefined;
	// Native explicit stops override the active theme stops; wholly themed tail
	// rows remain inherited placeholders and do not add colors to the gradient.
	const rows = sectionRows(sheet, `${paint}Gradient`)
		.filter(
			(row) =>
				row.cells.size === 0 || ![...row.cells.values()].every((cell) => cell.value === 'Themed'),
		)
		.slice(0, 10);
	// Pure theme placeholders belong to the DrawingML theme resolver.
	if (!rows.some((row) => [...row.cells.values()].some((cell) => cell.value !== 'Themed')))
		return undefined;
	const reject = () => {
		report(
			`unsupported-saved-${paint.toLowerCase()}-gradient`,
			`The saved ${paint.toLowerCase()} gradient has incomplete or unsupported settings; its foreground color is used.`,
		);
		return undefined;
	};
	const angle = number(cells, `${paint}GradientAngle`, NaN);
	const direction = number(cells, `${paint}GradientDir`, NaN);
	const wrapped = ((angle % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
	// Native raster paint uses physical projection. Stroke descriptors retain
	// normalized angles until the renderer supplies the physical line-width margin.
	const quarter = Math.round(wrapped / (Math.PI / 2));
	const orthogonal = Math.abs(wrapped - quarter * (Math.PI / 2)) < 1e-12;
	if (
		!Number.isFinite(width) ||
		!Number.isFinite(height) ||
		width < 0 ||
		height < 0 ||
		(paint === 'Fill' ? width === 0 || height === 0 : width === 0 && height === 0) ||
		(direction === 0 && !Number.isFinite(angle)) ||
		!Number.isInteger(direction) ||
		direction < 0 ||
		direction > 13 ||
		(paint === 'Line' && direction !== 0) ||
		number(cells, 'RotateGradientWithShape', NaN) !== 1 ||
		number(cells, 'UseGroupGradient', NaN) !== 0
	)
		return reject();
	// Visio uses only the first ten active stop rows. Do not sort stop positions:
	// coincident stops are meaningful and decreasing positions are malformed.
	if (rows.length < 2) return reject();
	const stops: VisioLinearGradient['stops'] = [];
	for (const row of rows) {
		const offset = number(row.cells, 'GradientStopPosition', NaN);
		const transparency = number(row.cells, 'GradientStopColorTrans', NaN);
		const value = row.cells.get('GradientStopColor')?.value;
		if (
			!Number.isFinite(offset) ||
			offset < 0 ||
			offset > 1 ||
			offset < (stops.at(-1)?.offset ?? 0) ||
			!Number.isFinite(transparency) ||
			transparency < 0 ||
			transparency > 1 ||
			!value ||
			value === 'Themed'
		)
			return reject();
		const color = resolveColor(row.cells);
		if (!color) return reject();
		stops.push({ offset, color, opacity: 1 - transparency });
	}
	const gradient: VisioFillGradient | undefined =
		direction === 13
			? pathFillGradient(geometry, width, height, stops)
			: direction >= 8
				? regionFillGradient(direction, stops)
				: direction !== 0
					? radialFillGradient(direction, stops, [width, height])
					: {
							type: 'linear',
							...((orthogonal && width > 0 && height > 0) || paint === 'Fill'
								? linearGradientEndpoints(
										width,
										height,
										orthogonal ? (quarter % 4) * 90 * 60_000 : (wrapped * 180 * 60_000) / Math.PI,
									)
								: {
										start: [0, 1] as const,
										end: [1, 1] as const,
										boundingBoxAngle: -Math.round((wrapped * 180 * 1e10) / Math.PI) / 1e10,
									}),
							stops,
						};
	if (!gradient) return reject();
	if (canUseVisioSigmaInterpolation(gradient)) gradient.interpolation = 'sigma-gamma22';
	return gradient;
}

export { savedShapeGradient as savedFillGradient };
