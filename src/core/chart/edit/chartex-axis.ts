// ChartEx axes have their own scaling and tick-label semantics; DrawingML formatting is shared.
import type { ChartExSpace, ChartExAxis } from '../ex/model';
import { editChartExSpace, setChartExTitle } from './chartex';
export interface ChartExAxisEdit {
	min?: number | null;
	max?: number | null;
	majorUnit?: number | null;
	minorUnit?: number | null;
	visible?: boolean;
	title?: string | null;
	tickLabels?: boolean;
	majorGridlines?: boolean;
	minorGridlines?: boolean;
	majorTickMark?: string | null;
	minorTickMark?: string | null;
	numberFormat?: ChartExAxis['numberFormat'] | null;
	units?: string | null;
}
export function editChartExAxis(
	space: ChartExSpace,
	index: number,
	edit: ChartExAxisEdit,
): ChartExSpace {
	if (!Number.isSafeInteger(index) || !space.plotArea.axes[index])
		throw new RangeError('No ChartEx axis at index.');
	for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) {
		const value = edit[key];
		if (
			value !== undefined &&
			value !== null &&
			(!Number.isFinite(value) || ((key === 'majorUnit' || key === 'minorUnit') && value <= 0))
		)
			throw new RangeError('Invalid ChartEx axis scaling.');
	}
	return editChartExSpace(space, (draft) => {
		const axis = draft.plotArea.axes[index]!;
		for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const)
			if (edit[key] !== undefined) {
				if (axis.scaling?.kind !== 'val')
					throw new RangeError('Only a value axis supports numeric scaling.');
				if (edit[key] === null) delete axis.scaling[key];
				else axis.scaling[key] = edit[key];
			}
		if (
			axis.scaling?.kind === 'val' &&
			typeof axis.scaling.min === 'number' &&
			typeof axis.scaling.max === 'number' &&
			axis.scaling.min >= axis.scaling.max
		)
			throw new RangeError('Axis maximum must exceed its minimum.');
		if (edit.visible !== undefined) axis.hidden = !edit.visible;
		if (edit.tickLabels !== undefined) {
			if (edit.tickLabels) axis.tickLabels ??= {};
			else delete axis.tickLabels;
		}
		for (const key of ['majorGridlines', 'minorGridlines'] as const)
			if (edit[key] !== undefined) {
				if (edit[key]) axis[key] ??= {};
				else delete axis[key];
			}
		for (const [key, target] of [
			['majorTickMark', 'majorTickMarks'],
			['minorTickMark', 'minorTickMarks'],
		] as const)
			if (edit[key] !== undefined) {
				const value = edit[key];
				if (value === null) delete axis[target];
				else {
					if (!['in', 'out', 'cross', 'none'].includes(value))
						throw new RangeError('Invalid tick mark.');
					axis[target] = { type: value };
				}
			}
		if (edit.title !== undefined) {
			const { title: _chartTitle, ...base } = draft;
			const title = setChartExTitle(
				{ ...base, ...(axis.title ? { title: axis.title } : {}) },
				edit.title ?? undefined,
			).title;
			if (title) axis.title = title;
			else delete axis.title;
		}
		if (edit.numberFormat !== undefined) {
			if (edit.numberFormat === null) delete axis.numberFormat;
			else axis.numberFormat = structuredClone(edit.numberFormat);
		}
		if (edit.units !== undefined) {
			if (edit.units === null) delete axis.units;
			else {
				if (
					![
						'hundreds',
						'thousands',
						'tenThousands',
						'hundredThousands',
						'millions',
						'tenMillions',
						'hundredMillions',
						'billions',
						'trillions',
						'percentage',
					].includes(edit.units)
				)
					throw new RangeError('Unsupported ChartEx display units.');
				axis.units = { ...axis.units, unit: edit.units };
			}
		}
	});
}
