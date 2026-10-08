// `c:chartSpace` writer: axes (`c:catAx`, `c:valAx`, `c:dateAx`, `c:serAx`). The shared part of
// the sequence is axId, scaling, delete, axPos, gridlines, title, numFmt, tick marks, tickLblPos,
// spPr, txPr, crossAx, crosses or crossesAt; the type-specific tail is written in one order that is
// consistent with all four types (crossBetween, auto, lblAlgn, lblOffset, baseTimeUnit, majorUnit,
// minorUnit, dispUnits, tickLblSkip, tickMarkSkip, noMultiLvlLbl, extLst).
import type { ChartAxis, ChartAxisKind } from './model';
import type { ChartShapeProperties } from './model-series';
import { numberFormatXml, shapePropertiesXml, textBodyXml } from './write-shape';
import { displayUnitsXml } from './write-chrome';
import { titleXml } from './write-text';
import { elementXml, raw, valXml, type ChartWriteContext } from './write-util';

const AXIS_ELEMENTS: Record<ChartAxisKind, string> = {
	cat: 'catAx',
	val: 'valAx',
	date: 'dateAx',
	ser: 'serAx',
};

const gridlinesXml = (
	context: ChartWriteContext,
	local: string,
	present: boolean,
	shape: ChartShapeProperties | undefined,
) => (present ? elementXml(local, shapePropertiesXml(context, shape)) : '');

/** One axis. */
export function axisXml(context: ChartWriteContext, axis: ChartAxis): string {
	const { scaling } = axis;
	return elementXml(
		AXIS_ELEMENTS[axis.kind],
		valXml('axId', axis.id) +
			elementXml(
				'scaling',
				valXml('logBase', scaling.logBase) +
					valXml('orientation', scaling.orientation) +
					valXml('max', scaling.max) +
					valXml('min', scaling.min),
			) +
			valXml('delete', axis.deleted) +
			valXml('axPos', axis.position) +
			gridlinesXml(context, 'majorGridlines', axis.majorGridlines, axis.majorGridlinesSpPr) +
			gridlinesXml(context, 'minorGridlines', axis.minorGridlines, axis.minorGridlinesSpPr) +
			titleXml(context, axis.title) +
			numberFormatXml(axis.numberFormat) +
			valXml('majorTickMark', axis.majorTickMark) +
			valXml('minorTickMark', axis.minorTickMark) +
			valXml('tickLblPos', axis.tickLabelPosition) +
			shapePropertiesXml(context, axis.spPr) +
			textBodyXml(context, axis.txPr) +
			valXml('crossAx', axis.crossAxisId) +
			valXml('crosses', axis.crosses) +
			valXml('crossesAt', axis.crossesAt) +
			valXml('crossBetween', axis.crossBetween) +
			valXml('auto', axis.auto) +
			valXml('lblAlgn', axis.labelAlign) +
			valXml('lblOffset', axis.labelOffset) +
			valXml('baseTimeUnit', axis.baseTimeUnit) +
			valXml('majorUnit', axis.majorUnit) +
			valXml('minorUnit', axis.minorUnit) +
			displayUnitsXml(context, axis.displayUnits) +
			valXml('tickLblSkip', axis.tickLabelSkip) +
			valXml('tickMarkSkip', axis.tickMarkSkip) +
			valXml('noMultiLvlLbl', axis.noMultiLevelLabels) +
			raw(context, axis.extLst),
	);
}
