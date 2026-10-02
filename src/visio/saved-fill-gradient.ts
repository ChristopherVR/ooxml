import type { VisioLinearGradient } from './model.js';
import { number, sectionRows, type Cells, type Report, type Sheet } from './sheet.js';
import { linearGradientEndpoints } from './theme-gradient.js';

/**
 * Saved ShapeSheet gradients use radians and normalized [0,1] stop values.
 * Only complete local, shape-rotating linear caches are accepted. Theme and
 * root-style substitution happen before this function; missing caches are not
 * inferred from formulas, legacy pattern numbers, or an unrelated theme.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/fill-gradient-section
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/e19c498a-5277-4add-9953-8b85cb2af250
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/64030657-ec1c-4bcd-adbb-701bb24004fa
 */
export function savedFillGradient(
	sheet: Sheet,
	width: number,
	height: number,
	resolveColor: (cells: Cells) => string,
	report: Report,
): VisioLinearGradient | undefined {
	const cells = sheet.cells;
	if (number(cells, 'FillGradientEnabled', NaN) !== 1) return undefined;
	const rows = sectionRows(sheet, 'FillGradient').slice(0, 10);
	// Pure theme placeholders belong to the DrawingML theme resolver.
	if (!rows.some((row) => [...row.cells.values()].some((cell) => cell.value !== 'Themed')))
		return undefined;
	const reject = () => {
		report(
			'unsupported-saved-fill-gradient',
			'The saved fill gradient has incomplete or unsupported settings; its foreground color is used.',
		);
		return undefined;
	};
	const angle = number(cells, 'FillGradientAngle', NaN);
	const wrapped = ((angle % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
	// Horizontal directions are sign-independent. Tolerate saved decimal rounding
	// (for example 3.14159265358979); other angles still need orientation evidence.
	const horizontal =
		Math.min(wrapped, 2 * Math.PI - wrapped) < 1e-12
			? 0
			: Math.abs(wrapped - Math.PI) < 1e-12
				? Math.PI
				: undefined;
	if (
		!Number.isFinite(width) ||
		!Number.isFinite(height) ||
		width <= 0 ||
		height <= 0 ||
		horizontal === undefined ||
		number(cells, 'FillGradientDir', NaN) !== 0 ||
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
	return {
		type: 'linear',
		...linearGradientEndpoints(width, height, ((horizontal * 180) / Math.PI) * 60_000),
		stops,
	};
}
