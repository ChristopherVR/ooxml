/** A small Office-shaped column chart part for the chart Format pane tests. */
import { parseChartSpace, type ChartSpace } from 'ooxml-core/chart';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const ser = (i: number, name: string, color: string) =>
	`<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${name}</c:v></c:tx>` +
	`<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>` +
	`<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>` +
	`<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>`;

const XML =
	`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart>` +
	`<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
	`<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
	ser(0, 'East', '4472C4') +
	ser(1, 'West', 'ED7D31') +
	`<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
	`<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>` +
	`<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>` +
	`</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>`;

export function columnSpace(): ChartSpace {
	const { chartSpace, issues } = parseChartSpace(XML);
	if (issues.length) throw new Error(issues.map((issue) => issue.message).join('\n'));
	return chartSpace;
}
