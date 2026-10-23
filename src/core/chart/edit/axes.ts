// Axis edits: visibility, title, scaling (bounds, units, logarithmic, reversed), number format, tick
// marks and labels, crossing, display units and gridlines. Field semantics follow PowerPoint's
// `setChartAxis`, `setChartAxisLogScale` and `setChartAxisGridlineStyle` (omit a field to keep it,
// `null` to return to automatic); unlike PowerPoint an axis that is not in the part is not created,
// because no chart group would reference it.
import type { ChartAxis, ChartAxisKind, ChartSpace } from '../model';
import type { ChartShapeProperties } from '../model-series';
import { assignOrDelete, editChartSpace } from './draft';
import { chartTitleWithText } from './title';

/** The first axis of a kind (PowerPoint's addressing) or the axis with an id. */
export type ChartAxisTarget = ChartAxisKind | { id: number };

export interface ChartAxisEdit {
	/** `c:delete` is the inverse. */
	visible?: boolean;
	/** Axis title text; `null` or `''` removes the title. */
	title?: string | null;
	min?: number | null;
	max?: number | null;
	majorUnit?: number | null;
	minorUnit?: number | null;
	/** Date axis: the unit of the major and minor units (`days`, `months`, `years`). */
	majorTimeUnit?: string | null;
	minorTimeUnit?: string | null;
	/** Logarithmic scale base (2..1000); `null` returns to a linear scale. */
	logBase?: number | null;
	/** Categories or values in reverse order (`c:orientation val="maxMin"`). */
	reverse?: boolean;
	/** A format code not linked to the source; `''` removes `c:numFmt` (the axis follows its data). */
	numberFormat?: string;
	majorTickMark?: 'cross' | 'in' | 'out' | 'none';
	minorTickMark?: 'cross' | 'in' | 'out' | 'none';
	tickLabelPosition?: 'high' | 'low' | 'nextTo' | 'none';
	/** Where the other axis crosses: `autoZero`, `min`, `max`, or a value. */
	crosses?: 'autoZero' | 'min' | 'max' | number;
	majorGridlines?: boolean;
	minorGridlines?: boolean;
	/** Value axis: a built-in unit (`thousands`, `millions`...) or a custom divisor; `null` clears. */
	displayUnits?: string | number | null;
}

/** The axis an edit targets; throws when the part has no such axis. */
export function chartAxisAt(space: ChartSpace, target: ChartAxisTarget): ChartAxis {
	const axis = space.plotArea.axes.find((item) =>
		typeof target === 'string' ? item.kind === target : item.id === target.id,
	);
	if (!axis) throw new RangeError(`No chart axis ${JSON.stringify(target)}`);
	return axis;
}

function setNumber(
	axis: ChartAxis,
	key: 'majorUnit' | 'minorUnit',
	value: number | null | undefined,
): void {
	if (value === undefined) return;
	if (value !== null && !(Number.isFinite(value) && value > 0))
		throw new RangeError(`Chart axis ${key} must be a positive number`);
	assignOrDelete(axis, key, value ?? undefined);
}

function applyAxis(axis: ChartAxis, edit: ChartAxisEdit): void {
	if (edit.visible !== undefined) axis.deleted = !edit.visible;
	if (edit.title !== undefined)
		assignOrDelete(
			axis,
			'title',
			edit.title ? chartTitleWithText(axis.title, edit.title) : undefined,
		);
	for (const key of ['min', 'max'] as const) {
		const value = edit[key];
		if (value === undefined) continue;
		if (value !== null && !Number.isFinite(value))
			throw new RangeError(`Chart axis ${key} must be a finite number`);
		assignOrDelete(axis.scaling, key, value ?? undefined);
	}
	const { min, max } = axis.scaling;
	if (min !== undefined && max !== undefined && min >= max)
		throw new RangeError('Chart axis minimum must be less than its maximum');
	setNumber(axis, 'majorUnit', edit.majorUnit);
	setNumber(axis, 'minorUnit', edit.minorUnit);
	if (edit.majorTimeUnit !== undefined)
		assignOrDelete(axis, 'majorTimeUnit', edit.majorTimeUnit ?? undefined);
	if (edit.minorTimeUnit !== undefined)
		assignOrDelete(axis, 'minorTimeUnit', edit.minorTimeUnit ?? undefined);
	if (edit.logBase !== undefined) {
		if (edit.logBase !== null && !(edit.logBase >= 2 && edit.logBase <= 1000))
			throw new RangeError('Chart axis log base must be from 2 to 1000');
		assignOrDelete(axis.scaling, 'logBase', edit.logBase ?? undefined);
	}
	if (edit.reverse !== undefined) axis.scaling.orientation = edit.reverse ? 'maxMin' : 'minMax';
	if (edit.numberFormat !== undefined)
		assignOrDelete(
			axis,
			'numberFormat',
			edit.numberFormat ? { formatCode: edit.numberFormat, sourceLinked: false } : undefined,
		);
	if (edit.majorTickMark !== undefined) axis.majorTickMark = edit.majorTickMark;
	if (edit.minorTickMark !== undefined) axis.minorTickMark = edit.minorTickMark;
	if (edit.tickLabelPosition !== undefined) axis.tickLabelPosition = edit.tickLabelPosition;
	if (edit.crosses !== undefined) {
		if (typeof edit.crosses === 'number') {
			delete axis.crosses;
			axis.crossesAt = edit.crosses;
		} else {
			delete axis.crossesAt;
			axis.crosses = edit.crosses;
		}
	}
	if (edit.majorGridlines !== undefined) {
		axis.majorGridlines = edit.majorGridlines;
		if (!edit.majorGridlines) delete axis.majorGridlinesSpPr;
	}
	if (edit.minorGridlines !== undefined) {
		axis.minorGridlines = edit.minorGridlines;
		if (!edit.minorGridlines) delete axis.minorGridlinesSpPr;
	}
	if (edit.displayUnits !== undefined) {
		if (axis.kind !== 'val') throw new RangeError('Only a value axis has display units');
		const units = edit.displayUnits;
		assignOrDelete(
			axis,
			'displayUnits',
			units === null
				? undefined
				: typeof units === 'number'
					? { customUnit: units }
					: { builtInUnit: units },
		);
	}
}

/** Edits one axis. */
export function setChartAxis(
	space: ChartSpace,
	target: ChartAxisTarget,
	edit: ChartAxisEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => applyAxis(chartAxisAt(draft, target), edit));
}

/**
 * Sets the outline of an axis's major or minor gridlines (colour, width, dash), keeping the rest
 * of their shape; styling gridlines turns them on, as in PowerPoint. `null` removes the styling.
 */
export function setChartGridlineShape(
	space: ChartSpace,
	target: ChartAxisTarget,
	which: 'major' | 'minor',
	shape: Pick<ChartShapeProperties, 'line' | 'fill'> | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const axis = chartAxisAt(draft, target);
		const key = which === 'major' ? 'majorGridlinesSpPr' : 'minorGridlinesSpPr';
		if (shape === null) {
			delete axis[key];
			return;
		}
		const next: ChartShapeProperties = { ...axis[key] };
		if (shape.line !== undefined) next.line = structuredClone(shape.line);
		if (shape.fill !== undefined) next.fill = structuredClone(shape.fill);
		axis[key] = next;
		axis[which === 'major' ? 'majorGridlines' : 'minorGridlines'] = true;
	});
}
