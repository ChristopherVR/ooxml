import { chartTitleRichXml, newChartTitle } from '../../chart/edit/title';

/** A new chart title (`c:title`) as XML, for patching a kept part. */
export function chartTitleXml(title: string): string {
	return `<c:title><c:tx><c:rich>${chartTitleRichXml(title)}</c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

/** A new chart title in the neutral model, written as {@link chartTitleXml} writes it. */
export const chartTitle = newChartTitle;
