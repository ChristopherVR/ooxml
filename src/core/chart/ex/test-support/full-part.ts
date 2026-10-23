// A chartex part using every element and attribute the neutral model carries, in schema order,
// shared by the chartex round-trip test and the pptx parity harness.
export const NS =
	'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const fill = (hex: string) =>
	`<cx:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></cx:spPr>`;
const ext = '<cx:extLst><a:ext uri="{X}"><a:x val="1"/></a:ext></cx:extLst>';

/** One part using every element and attribute the model carries, in schema order. */
export const FULL =
	`${DECLARATION}<cx:chartSpace ${NS}><cx:chartData><cx:externalData r:id="rId3" cx:autoUpdate="1"/>` +
	'<cx:data id="0"><cx:strDim type="cat"><cx:f dir="row">Sheet1!$A$2:$B$4</cx:f><cx:nf>Sheet1!$A$1</cx:nf>' +
	'<cx:lvl ptCount="3" name="Leaf"><cx:pt idx="0">a</cx:pt><cx:pt idx="2"> c </cx:pt></cx:lvl>' +
	'<cx:lvl ptCount="3"><cx:pt idx="0">P</cx:pt></cx:lvl></cx:strDim>' +
	'<cx:numDim type="val"><cx:f>Sheet1!$C$2:$C$4</cx:f><cx:lvl ptCount="3" formatCode="0.0">' +
	`<cx:pt idx="0">1.50</cx:pt><cx:pt idx="1">-2</cx:pt>${ext}</cx:lvl>${ext}</cx:numDim>${ext}</cx:data>` +
	'<cx:data id="1"><cx:numDim type="colorVal"><cx:lvl ptCount="1"><cx:pt idx="0">7</cx:pt></cx:lvl></cx:numDim>' +
	'<cx:strDim type="entityId"><cx:lvl ptCount="1"><cx:pt idx="0">US</cx:pt></cx:lvl></cx:strDim></cx:data></cx:chartData>' +
	'<cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Title</a:t></a:r></a:p></cx:rich></cx:tx>' +
	`${fill('112233')}<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>${ext}</cx:title>` +
	`<cx:plotArea><cx:plotAreaRegion><cx:plotSurface>${fill('EEEEEE')}</cx:plotSurface>` +
	'<cx:series layoutId="clusteredColumn" hidden="0" uniqueId="{A}" formatIdx="0"><cx:tx><cx:txData><cx:f>Sheet1!$C$1</cx:f><cx:v>Sales</cx:v></cx:txData></cx:tx>' +
	`${fill('4472C4')}<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>` +
	`<cx:dataPt idx="1">${fill('FF0000')}</cx:dataPt>` +
	'<cx:dataLabels pos="outEnd"><cx:numFmt formatCode="#,##0" sourceLinked="0"/><cx:visibility seriesName="0" categoryName="1" value="1"/><cx:separator>; </cx:separator>' +
	'<cx:dataLabel idx="2" pos="ctr"><cx:visibility seriesName="1" categoryName="0" value="0"/></cx:dataLabel><cx:dataLabelHidden idx="0"/></cx:dataLabels>' +
	'<cx:dataId val="0"/><cx:layoutPr><cx:binning intervalClosed="l" underflow="auto" overflow="9.5"><cx:binSize val="2.5"/></cx:binning></cx:layoutPr>' +
	'<cx:axisId val="0"/><cx:axisId val="1"/></cx:series>' +
	'<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="{B}"><cx:axisId val="2"/></cx:series>' +
	'<cx:series layoutId="regionMap" uniqueId="{C}"><cx:valueColors><cx:minColor><a:srgbClr val="FFFFFF"/></cx:minColor></cx:valueColors>' +
	'<cx:valueColorPositions count="2"><cx:min><cx:extremeValue/></cx:min></cx:valueColorPositions><cx:dataId val="1"/>' +
	'<cx:layoutPr><cx:regionLabelLayout val="bestFitOnly"/><cx:geography projectionType="mercator" viewedRegionType="countryRegion" cultureLanguage="en-US" cultureRegion="US" attribution="Powered by Bing">' +
	'<cx:geoCache provider="{P}"><cx:binary>QUJD</cx:binary></cx:geoCache></cx:geography></cx:layoutPr></cx:series>' +
	'<cx:series layoutId="boxWhisker"><cx:layoutPr><cx:visibility meanLine="1" meanMarker="0" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr></cx:series>' +
	'<cx:series layoutId="waterfall"><cx:layoutPr><cx:visibility connectorLines="0"/><cx:subtotals><cx:idx val="0"/><cx:idx val="3"/></cx:subtotals></cx:layoutPr></cx:series>' +
	'<cx:series layoutId="treemap"><cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr></cx:series>' +
	'<cx:series layoutId="clusteredColumn"><cx:layoutPr><cx:aggregation/></cx:layoutPr></cx:series></cx:plotAreaRegion>' +
	'<cx:axis id="0"><cx:catScaling gapWidth="auto"/><cx:title><cx:tx><cx:txData><cx:v>Bins</cx:v></cx:txData></cx:tx></cx:title><cx:majorTickMarks type="out"/><cx:tickLabels/></cx:axis>' +
	'<cx:axis id="1" hidden="1"><cx:valScaling max="100" min="auto" majorUnit="20" minorUnit="auto"/>' +
	'<cx:units unit="thousands"><cx:unitsLabel><cx:tx><cx:txData><cx:v>Thousands</cx:v></cx:txData></cx:tx></cx:unitsLabel></cx:units>' +
	`<cx:majorGridlines>${fill('D9D9D9')}</cx:majorGridlines><cx:minorGridlines/><cx:majorTickMarks/><cx:minorTickMarks type="none"/>` +
	`<cx:tickLabels>${ext}</cx:tickLabels><cx:numFmt formatCode="0%" sourceLinked="1"/>${fill('000000')}` +
	'<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr></cx:axis>' +
	'<cx:axis id="2"><cx:valScaling/></cx:axis>' +
	`${fill('F2F2F2')}</cx:plotArea><cx:legend pos="r" align="min" overlay="1">${fill('FFFFFF')}</cx:legend>${ext}</cx:chart>` +
	`${fill('FFFFFF')}<cx:clrMapOvr bg1="lt1" tx1="dk1"/><cx:fmtOvrs><cx:fmtOvr idx="0"/></cx:fmtOvrs>` +
	`<cx:printSettings><cx:pageMargins l="0.7"/></cx:printSettings>${ext}</cx:chartSpace>`;

