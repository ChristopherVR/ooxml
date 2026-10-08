/**
 * pptx chart axes (`PptxChartAxisFormatting`) from the neutral axes (`ChartAxis`): ids, position,
 * deletion, scaling, gridlines, number format, title text, tick marks and labels, crossing, units
 * and display units come from the neutral model; the resolved styles (shape properties, fonts,
 * gridline and display-unit label styling, date-axis major and minor time units) stay with the
 * object-tree reader (`parseChartAxes`), whose result is passed in as `styled`. Every value is
 * checked against its schema type, as the object-tree reader checks it. The list keeps the pptx
 * order: grouped by axis element, each group in document order.
 *
 * @module chart-from-neutral-axes
 */
import type { ChartAxis, ChartTitle } from '../../../chart/index';
import type { PptxChartAxisFormatting } from '../types';
import { setOrRemove } from './chart-from-neutral-labels';

type Axis = PptxChartAxisFormatting;

const oneOf = <T extends string>(values: readonly T[], value: string | undefined): T | undefined =>
	value !== undefined && (values as readonly string[]).includes(value) ? (value as T) : undefined;
const finite = (value: number | undefined) =>
	value !== undefined && Number.isFinite(value) ? value : undefined;
const positiveInteger = (value: number | undefined) =>
	value !== undefined && Number.isInteger(value) && value > 0 ? value : undefined;

const TICK_MARKS = ['cross', 'in', 'none', 'out'] as const;
const TIME_UNITS = ['days', 'months', 'years'] as const;
const BUILT_IN_UNITS = [
	'hundreds',
	'thousands',
	'tenThousands',
	'hundredThousands',
	'millions',
	'tenMillions',
	'hundredMillions',
	'billions',
	'trillions',
] as const;

/** The concatenated run text of an axis title (no paragraph breaks), when it has runs. */
function titleText(title: ChartTitle | undefined): string | undefined {
	const runs = title?.tx?.rich?.paragraphs.flatMap((paragraph) => paragraph.runs) ?? [];
	return runs.length > 0 ? runs.map((run) => run.text).join('') : undefined;
}

function scalingFromNeutral(axis: ChartAxis, out: Axis): void {
	const { scaling } = axis;
	setOrRemove(out, 'min', finite(scaling.min));
	setOrRemove(out, 'max', finite(scaling.max));
	const logBase =
		scaling.logBase !== undefined && scaling.logBase > 0 ? scaling.logBase : undefined;
	setOrRemove(out, 'logBase', logBase);
	setOrRemove(out, 'logScale', logBase !== undefined ? true : undefined);
	setOrRemove(out, 'orientation', oneOf(['minMax', 'maxMin'], scaling.orientation));
}

function labelsFromNeutral(axis: ChartAxis, out: Axis): void {
	const type = out.axisType;
	setOrRemove(out, 'majorTickMark', oneOf(TICK_MARKS, axis.majorTickMark));
	setOrRemove(out, 'minorTickMark', oneOf(TICK_MARKS, axis.minorTickMark));
	setOrRemove(out, 'tickLblPos', oneOf(['high', 'low', 'nextTo', 'none'], axis.tickLabelPosition));
	setOrRemove(out, 'crosses', oneOf(['autoZero', 'min', 'max'], axis.crosses));
	setOrRemove(out, 'crossesAt', finite(axis.crossesAt));
	const valueAxis = type === 'valAx';
	const categoryOrDate = type === 'catAx' || type === 'dateAx';
	const skips = type === 'catAx' || type === 'serAx';
	setOrRemove(
		out,
		'crossBetween',
		valueAxis ? oneOf(['between', 'midCat'], axis.crossBetween) : undefined,
	);
	setOrRemove(out, 'auto', categoryOrDate ? axis.auto : undefined);
	const offset = axis.labelOffset;
	setOrRemove(
		out,
		'labelOffset',
		categoryOrDate && offset !== undefined && offset >= 0 && offset <= 1000 ? offset : undefined,
	);
	setOrRemove(out, 'tickLabelSkip', skips ? positiveInteger(axis.tickLabelSkip) : undefined);
	setOrRemove(out, 'tickMarkSkip', skips ? positiveInteger(axis.tickMarkSkip) : undefined);
	setOrRemove(
		out,
		'labelAlignment',
		type === 'catAx' ? oneOf(['ctr', 'l', 'r'], axis.labelAlign) : undefined,
	);
	setOrRemove(out, 'noMultiLevelLabels', type === 'catAx' ? axis.noMultiLevelLabels : undefined);
}

function axisFromNeutral(axis: ChartAxis, out: Axis): Axis {
	setOrRemove(out, 'axPos', oneOf(['b', 'l', 'r', 't'], axis.position));
	const formatCode = axis.numberFormat?.formatCode.trim();
	setOrRemove(
		out,
		'numFmt',
		formatCode ? { formatCode, sourceLinked: axis.numberFormat?.sourceLinked === true } : undefined,
	);
	setOrRemove(out, 'titleText', titleText(axis.title));
	setOrRemove(out, 'axisId', axis.id);
	setOrRemove(out, 'crossAxisId', axis.crossAxisId);
	setOrRemove(out, 'deleted', axis.deleted === true ? true : undefined);
	scalingFromNeutral(axis, out);
	labelsFromNeutral(axis, out);
	if (!axis.majorGridlines) {
		delete out.majorGridlines;
		delete out.majorGridlinesSpPr;
	} else out.majorGridlines = true;
	if (!axis.minorGridlines) {
		delete out.minorGridlines;
		delete out.minorGridlinesSpPr;
	} else out.minorGridlines = true;
	const units = axis.displayUnits;
	if (units) {
		const custom = units.customUnit;
		const builtIn = oneOf(BUILT_IN_UNITS, units.builtInUnit?.trim());
		setOrRemove(out, 'displayUnits', custom ? 'custom' : builtIn);
		setOrRemove(out, 'displayUnitsValue', custom || undefined);
	}
	setOrRemove(out, 'majorUnit', finite(axis.majorUnit));
	setOrRemove(out, 'minorUnit', finite(axis.minorUnit));
	setOrRemove(
		out,
		'baseTimeUnit',
		out.axisType === 'dateAx' ? oneOf(TIME_UNITS, axis.baseTimeUnit) : undefined,
	);
	return out;
}

/**
 * The pptx axis list: the neutral axes grouped by element (in order of first appearance), each
 * filled onto the object-tree axis of the same element and position, which keeps its styles.
 */
export function axesFromNeutral(axes: ChartAxis[], styled: Axis[]): Axis[] {
	const kinds: ChartAxis['kind'][] = [];
	for (const axis of axes) if (!kinds.includes(axis.kind)) kinds.push(axis.kind);
	return kinds.flatMap((kind) => {
		const type = `${kind}Ax` as Axis['axisType'];
		const sameType = styled.filter((axis) => axis.axisType === type);
		return axes
			.filter((axis) => axis.kind === kind)
			.map((axis, at) => axisFromNeutral(axis, sameType[at] ?? { axisType: type }));
	});
}
