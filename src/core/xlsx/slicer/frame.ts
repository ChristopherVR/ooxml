import { NS, children, elements, first, type XmlElement } from '../../xml/index';
import type { DrawingAnchor, SlicerFrameObject } from '../model';

export const SLICER_FRAME_URI = 'http://schemas.microsoft.com/office/drawing/2010/slicer';
export const TIMELINE_FRAME_URI = 'http://schemas.microsoft.com/office/drawing/2012/timeslicer';

/** Reads a slicer or timeline frame (`mc:AlternateContent` > `xdr:graphicFrame`), or undefined. */
export function readSlicerFrame(
	content: XmlElement,
	anchor: DrawingAnchor,
	sourceXml: string,
): SlicerFrameObject | undefined {
	for (const choice of children(content, 'Choice', NS.mc)) {
		const frame = first(choice, 'graphicFrame', NS.xdr);
		const data = first(first(frame, 'graphic', NS.a), 'graphicData', NS.a);
		const uri = data?.getAttribute('uri');
		if (uri !== SLICER_FRAME_URI && uri !== TIMELINE_FRAME_URI) continue;
		const name = data ? elements(data)[0]?.getAttribute('name') : undefined;
		if (!name) return undefined;
		return uri === SLICER_FRAME_URI
			? { kind: 'slicer', anchor, name, sourceXml }
			: { kind: 'timeline', anchor, name, sourceXml };
	}
	return undefined;
}

const escape = (text: string) =>
	text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** A fresh frame for a slicer or timeline, with the "not supported" fallback Excel writes. */
export function slicerFrameXml(
	kind: 'slicer' | 'timeline',
	name: string,
	id: number,
	anchor: DrawingAnchor,
	box: { x: number; y: number; cx: number; cy: number },
): string {
	const slicer = kind === 'slicer';
	const uri = slicer ? SLICER_FRAME_URI : TIMELINE_FRAME_URI;
	const requires = slicer ? 'sle15' : 'tsle';
	const ns = slicer
		? 'xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer"'
		: 'xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer"';
	const tag = slicer
		? `<sle:slicer xmlns:sle="${uri}" name="${escape(name)}"/>`
		: `<tsle:timeslicer xmlns:tsle="${uri}" name="${escape(name)}"/>`;
	const note = slicer
		? 'This shape represents a table slicer. Table slicers are not supported in this version of Excel.'
		: 'Timeline: Works in Excel 2013 or higher. Do not move or resize.';
	const marker = (tag2: string, m: DrawingAnchor['from']) =>
		`<xdr:${tag2}><xdr:col>${m.col}</xdr:col><xdr:colOff>${Math.round(m.colOffset)}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${Math.round(m.rowOffset)}</xdr:rowOff></xdr:${tag2}>`;
	const to = anchor.to ?? anchor.from;
	return (
		`<xdr:twoCellAnchor editAs="absolute" xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${marker('from', anchor.from)}${marker('to', to)}` +
		`<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice ${ns} Requires="${requires}">` +
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escape(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
		`<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${tag}</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
		`<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
		`<xdr:spPr><a:xfrm><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.cx)}" cy="${Math.round(box.cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
		`<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>${escape(note)}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
	);
}
