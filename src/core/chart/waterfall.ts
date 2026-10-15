// Waterfall (ChartEx `waterfall` layout) arithmetic shared by every product that draws one:
// running-total steps with absolute subtotal bars, and the value range that contains them.
import { axisTargetIntervals, niceValueAxisBounds } from './axis-nice';

/** The subtotal marks of `cx:layoutPr/cx:subtotals`. */
export interface WaterfallOptions {
	/** Indexes of points that are absolute totals; absent means the last point only. */
	subtotalIndices?: readonly number[] | undefined;
	/** `cx:visibility/@connectorLines`; connectors are drawn unless this is `false`. */
	connectorLines?: boolean | undefined;
}

export interface WaterfallStep {
	sourceIndex: number;
	value: number;
	startValue: number;
	endValue: number;
	isSubtotal: boolean;
}

/** A resolved automatic value range. */
export interface WaterfallRange {
	min: number;
	max: number;
	span: number;
	majorUnit?: number;
}

/** Resolve ChartEx waterfall deltas and absolute subtotal bars. */
export function buildWaterfallSteps(
	values: ReadonlyArray<number>,
	options: Pick<WaterfallOptions, 'subtotalIndices'> | undefined,
): WaterfallStep[] {
	const subtotalIndices = new Set(
		options?.subtotalIndices ?? (values.length > 0 ? [values.length - 1] : []),
	);
	const steps: WaterfallStep[] = [];
	let runningTotal = 0;
	for (let sourceIndex = 0; sourceIndex < values.length; sourceIndex++) {
		const value = values[sourceIndex] ?? 0;
		const isSubtotal = subtotalIndices.has(sourceIndex);
		const startValue = isSubtotal ? 0 : runningTotal;
		const endValue = isSubtotal ? value : runningTotal + value;
		steps.push({ sourceIndex, value, startValue, endValue, isSubtotal });
		runningTotal = endValue;
	}
	return steps;
}

/**
 * A value range containing every cumulative and absolute bar endpoint.
 *
 * @param plotHeightPx The plot-area height, when known: it sets the automatic scale's target
 *   interval count (`axisTargetIntervals`).
 */
export function computeWaterfallRange(
	steps: ReadonlyArray<WaterfallStep>,
	plotHeightPx?: number,
): WaterfallRange {
	let dataMin = Number.POSITIVE_INFINITY;
	let dataMax = Number.NEGATIVE_INFINITY;
	for (const step of steps)
		for (const value of [step.startValue, step.endValue]) {
			if (value < dataMin) dataMin = value;
			if (value > dataMax) dataMax = value;
		}
	if (dataMin === Number.POSITIVE_INFINITY) return { min: 0, max: 1, span: 1 };
	const { min, max, majorUnit } = niceValueAxisBounds(
		dataMin,
		dataMax,
		plotHeightPx !== undefined ? axisTargetIntervals(plotHeightPx) : undefined,
	);
	return { min, max, span: Math.max(max - min, Number.EPSILON), majorUnit };
}
