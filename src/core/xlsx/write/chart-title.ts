import type { ChartTitle } from '../../chart/model-series';
import { chartTextBodyFromXml } from '../../chart/write-shape';
import { escapeText } from './xml-out';

/** Keep native rich-text property nodes so Excel can resolve inherited title fonts. */
const titleRichXml = (title: string) =>
	`<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeText(title)}</a:t></a:r></a:p>`;

/** A new chart title (`c:title`) as XML, for patching a kept part. */
export function chartTitleXml(title: string): string {
	return `<c:title><c:tx><c:rich>${titleRichXml(title)}</c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

/** A new chart title in the neutral model, written as {@link chartTitleXml} writes it. */
export function chartTitle(title: string): ChartTitle {
	const rich = chartTextBodyFromXml(titleRichXml(title));
	return { tx: { rich, text: rich.text }, text: rich.text, overlay: false };
}
