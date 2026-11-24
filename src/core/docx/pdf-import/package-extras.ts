/**
 * The parts of the reconstruction beyond plain paragraphs: footnotes (`word/footnotes.xml` with
 * Word's separator notes), footnote references, bookmarks and the text boxes that hold rotated
 * text.
 */
import { esc, NS, ROOT_NAMESPACES, XML_DECLARATION } from '../load/legacy-doc/xml';
import type { Fragment } from './lines';
import type { FloatingText } from './rotated';

const EMU = 12700;
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

export const FOOTNOTE_STYLES =
	'<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
	'<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';

/** The run of a footnote reference mark (a custom mark when the PDF's mark is not the number). */
export function noteReferenceRun(note: NonNullable<Fragment['note']>): string {
	const style = '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>';
	// A custom mark is the text that follows the reference, written as a run of its own.
	return (
		`<w:r>${style}<w:footnoteReference${note.custom ? ' w:customMarkFollows="1"' : ''} w:id="${note.id}"/></w:r>` +
		(note.custom ? `<w:r>${style}<w:t>${esc(note.mark)}</w:t></w:r>` : '')
	);
}

/** The mark that starts a note's own text: its number, or the custom mark as text. */
export function noteMarkRun(customMark?: string): string {
	const mark = customMark === undefined ? '<w:footnoteRef/>' : `<w:t>${esc(customMark)}</w:t>`;
	return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>${mark}</w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`;
}

/** `word/footnotes.xml`: the separator notes Word expects, then the notes' paragraphs. */
export function footnotesXml(notes: { id: number; body: string }[]): string {
	const separator = (type: string, id: number, mark: string) =>
		`<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${mark}/></w:r></w:p></w:footnote>`;
	return (
		`${XML_DECLARATION}<w:footnotes ${ROOT_NAMESPACES}>` +
		separator('separator', -1, 'separator') +
		separator('continuationSeparator', 0, 'continuationSeparator') +
		notes.map((note) => `<w:footnote w:id="${note.id}">${note.body}</w:footnote>`).join('') +
		'</w:footnotes>'
	);
}

/** Bookmark starts and ends around a paragraph's runs. */
export function bookmarkXml(names: readonly string[] | undefined, ids: { next: number }) {
	if (!names?.length) return { start: '', end: '' };
	const numbered = names.map((name) => ({ name, id: ids.next++ }));
	return {
		start: numbered
			.map((item) => `<w:bookmarkStart w:id="${item.id}" w:name="${esc(item.name)}"/>`)
			.join(''),
		end: numbered.map((item) => `<w:bookmarkEnd w:id="${item.id}"/>`).join(''),
	};
}

/**
 * A floating text box with `content` (its paragraphs), placed on the page where the PDF drew the
 * text: vertical text for 90 and 270 degrees, a rotated box otherwise.
 */
export function textBoxXml(float: FloatingText, content: string, id: number): string {
	const vertical = float.angle === 90 || float.angle === 270;
	const margin = 4;
	// The box before rotation: along the text, then across it.
	// Some room along the text, so a font measuring a little wider does not wrap the line.
	const along = float.length * 1.1 + 2 * margin;
	const across = float.thickness + 2 * margin;
	const [boxWidth, boxHeight] = vertical ? [across, along] : [along, across];
	const centreX = float.x + float.width / 2;
	const centreY = float.y + float.height / 2;
	const cx = Math.max(1, Math.round(boxWidth * EMU));
	const cy = Math.max(1, Math.round(boxHeight * EMU));
	const x = Math.round((centreX - boxWidth / 2) * EMU);
	const y = Math.round((centreY - boxHeight / 2) * EMU);
	const rotation = vertical ? 0 : Math.round(((360 - float.angle) % 360) * 60000);
	const vert = float.angle === 90 ? 'vert270' : float.angle === 270 ? 'vert' : 'horz';
	const inset = Math.round(margin * EMU);
	const graphic =
		`<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
		`<wps:spPr><a:xfrm${rotation ? ` rot="${rotation}"` : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
		`<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
		`<wps:bodyPr rot="0" vert="${vert}" wrap="square" lIns="${inset}" tIns="${inset}" rIns="${inset}" bIns="${inset}" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>`;
	return (
		`<w:r><mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="wps"><w:drawing>` +
		`<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id + 1}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
		`<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
		`<wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>${graphic}</wp:anchor>` +
		'</w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r>'
	);
}
