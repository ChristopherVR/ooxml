// A small Office-shaped chart part the edit tests start from: a clustered column chart with two
// series on a category and a value axis, a rich title and a legend.
import { NS } from '../../../xml/index';
import type { ChartSpace } from '../../model';
import { parseChartSpace } from '../../parse-space';

const ser = (i: number, name: string, values: number[], color: string) =>
	`<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${'BC'[i]}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
	`<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
	`<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt><c:pt idx="2"><c:v>East</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
	`<c:val><c:numRef><c:f>Sheet1!$${'BC'[i]}$2:$${'BC'[i]}$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${values
		.map((v, p) => `<c:pt idx="${p}"><c:v>${v}</c:v></c:pt>`)
		.join('')}</c:numCache></c:numRef></c:val></c:ser>`;

export const COLUMN_CHART_XML =
	`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/><c:chart>` +
	`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Sales</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> 2026</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
	`<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
	ser(0, 'Plan', [10, 20, 30], '4472C4') +
	ser(1, 'Actual', [12, 18, 33], 'ED7D31') +
	`<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="11"/><c:axId val="22"/></c:barChart>` +
	`<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="22"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
	`<c:valAx><c:axId val="22"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="11"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` +
	`</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;

/** The parsed column chart. */
export function columnChart(): ChartSpace {
	const { chartSpace, issues } = parseChartSpace(COLUMN_CHART_XML);
	if (issues.length) throw new Error(issues.map((issue) => issue.message).join('\n'));
	return chartSpace;
}
