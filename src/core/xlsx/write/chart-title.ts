import { escapeText } from './xml-out';

/** Keep native rich-text property nodes so Excel can resolve inherited title fonts. */
export function chartTitleXml(title: string): string {
	return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeText(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}
