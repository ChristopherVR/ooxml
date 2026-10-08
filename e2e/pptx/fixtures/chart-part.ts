/**
 * Package plumbing for the chart fixture generators: the core save pipeline
 * has no path to author a brand-new chart part (see `generate-chart-fixture.ts`),
 * so a generator saves a base deck and then injects the chart part, its slide
 * relationship, the slide's `p:graphicFrame` and the content-type override.
 *
 * @module e2e/fixtures/chart-part
 */

const EMU_PER_PX = 9525;

/** The frame box in CSS pixels on a 960x540 slide. */
export interface ChartFrameBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

const DEFAULT_BOX: ChartFrameBox = { x: 60, y: 60, width: 840, height: 420 };

/** A `p:graphicFrame` referencing the chart part behind relationship `rId`. */
export function chartGraphicFrameXml(rId: string, box: ChartFrameBox = DEFAULT_BOX): string {
	return (
		`<p:graphicFrame><p:nvGraphicFramePr>` +
		`<p:cNvPr id="101" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
		`<p:xfrm><a:off x="${box.x * EMU_PER_PX}" y="${box.y * EMU_PER_PX}"/>` +
		`<a:ext cx="${box.width * EMU_PER_PX}" cy="${box.height * EMU_PER_PX}"/></p:xfrm>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
		`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/>` +
		`</a:graphicData></a:graphic></p:graphicFrame>`
	);
}

/** Append `frameXml` as the last child of the slide's shape tree. */
export function injectGraphicFrame(slideXml: string, frameXml: string): string {
	const marker = '</p:spTree>';
	const at = slideXml.lastIndexOf(marker);
	if (at < 0) {
		throw new Error('slide XML missing </p:spTree>');
	}
	return slideXml.slice(0, at) + frameXml + slideXml.slice(at);
}

/** Add a chart relationship to a slide's `.rels`, returning the new XML and its id. */
export function addChartRel(relsXml: string, target: string): { xml: string; rId: string } {
	const ids = [...relsXml.matchAll(/Id="rId(?<n>\d+)"/gu)].map((m) =>
		Number.parseInt(m.groups?.n ?? '0', 10),
	);
	const rId = `rId${(ids.length > 0 ? Math.max(...ids) : 0) + 1}`;
	const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
	const rel = `<Relationship Id="${rId}" Type="${relType}" Target="${target}"/>`;
	return { xml: relsXml.replace('</Relationships>', `${rel}</Relationships>`), rId };
}

/** Register `partName` (no leading slash) as a chart part in `[Content_Types].xml`. */
export function addContentTypeOverride(ctXml: string, partName: string): string {
	const contentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
	const override = `<Override PartName="/${partName}" ContentType="${contentType}"/>`;
	return ctXml.replace('</Types>', `${override}</Types>`);
}
