import type { ChartTextEdit } from './text';

/** Shared validation for chart targets and standalone rich/title/legend-entry text bodies. */
export function validateChartTextEdit(edit: ChartTextEdit): void {
	if (
		edit.rotation !== undefined &&
		edit.rotation !== null &&
		(!Number.isFinite(edit.rotation) || Math.abs(edit.rotation) > 360)
	)
		throw new RangeError('Invalid chart text rotation');
	if (
		edit.sizePt !== undefined &&
		edit.sizePt !== null &&
		(!Number.isFinite(edit.sizePt) || edit.sizePt <= 0 || edit.sizePt > 4000)
	)
		throw new RangeError('Invalid chart text size');
	if (
		edit.anchor !== undefined &&
		edit.anchor !== null &&
		!['t', 'ctr', 'b', 'just', 'dist'].includes(edit.anchor)
	)
		throw new RangeError('Invalid chart text anchor');
	if (
		edit.direction !== undefined &&
		edit.direction !== null &&
		![
			'horz',
			'vert',
			'vert270',
			'wordArtVert',
			'eaVert',
			'mongolianVert',
			'wordArtVertRtl',
		].includes(edit.direction)
	)
		throw new RangeError('Invalid chart text direction');
	for (const key of ['fillTransparency', 'lineTransparency'] as const)
		if (
			edit[key] !== undefined &&
			(!Number.isFinite(edit[key]) || edit[key] < 0 || edit[key] > 100)
		)
			throw new RangeError('Invalid chart text transparency');
	for (const key of ['typeface', 'eastAsiaTypeface'] as const)
		if (edit[key] !== undefined && edit[key] !== null && !edit[key].trim())
			throw new RangeError('Invalid chart font family');
}
