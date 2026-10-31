import { NS } from '../../xml/index';
import { formatAddress } from '../address';
import type { Comment } from '../model';
import { XML_HEADER, escapeText, tElement } from './xml-out';
import type { PersonRegistry } from './comment-persons';
import { ThreadedCommentsWriter, type ThreadedCommentOptions } from './threaded-comments';
export { PersonRegistry } from './comment-persons';
export { guidFrom } from './comment-ids';

const THREADED_NOTE =
	'[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ';

export interface CommentParts {
	comments: string;
	vml: string;
	/** Present when at least one comment is threaded (has a `replies` array). */
	threaded?: string;
}

function vmlShape(comment: Comment, index: number, id: number): string {
	const { row, col } = comment.address;
	const anchor = `${col + 1}, 15, ${Math.max(0, row - 1)}, 10, ${col + 3}, 15, ${row + 3}, 4`;
	return (
		`<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
		'<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
		'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
		`<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`
	);
}

/** Where the comment shapes of a VML drawing go (see `VmlIdAllocator`). */
export interface VmlLayout {
	/** The drawing's `o:idmap` blocks. */
	blocks: readonly number[];
	/** One shape id per comment, in sheet order. */
	ids: readonly number[];
	/** A source drawing (comment shapes removed) whose other shapes, such as form controls, stay. */
	base?: string;
}

const NOTE_SHAPETYPE =
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

function vmlXml(comments: readonly Comment[], layout: VmlLayout): string {
	const shapes = comments
		.map((comment, index) => vmlShape(comment, index, layout.ids[index] ?? 0))
		.join('');
	const data = layout.blocks.join(',');
	const base = layout.base;
	if (base && /<\/xml>\s*$/.test(base)) {
		let xml = /<o:idmap\b/.test(base)
			? base.replace(/(<o:idmap\b[^>]*\bdata=")[^"]*"/, `$1${data}"`)
			: base.replace(
					/(<xml\b[^>]*>)/,
					`$1<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${data}"/></o:shapelayout>`,
				);
		if (!/<v:shapetype\b[^>]*\bid="_x0000_t202"/.test(xml))
			xml = xml.replace(/<\/xml>\s*$/, `${NOTE_SHAPETYPE}</xml>`);
		return xml.replace(/<\/xml>\s*$/, `${shapes}</xml>`);
	}
	return (
		`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
		`<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${data}"/></o:shapelayout>` +
		NOTE_SHAPETYPE +
		shapes +
		'</xml>'
	);
}

/**
 * The comments part, the legacy VML drawing that makes Excel show them, and (for threaded
 * comments) the threaded comments part. `sheetNumber` keeps VML shape ids unique per sheet
 * unless `layout` gives the ids (and any source drawing to keep).
 */
export function commentParts(
	comments: readonly Comment[],
	sheetNumber: number,
	persons: PersonRegistry,
	sheetName: string,
	layout?: VmlLayout,
	threadOptions?: ThreadedCommentOptions,
): CommentParts {
	const authors: string[] = [];
	const authorId = (name: string) => {
		let id = authors.indexOf(name);
		if (id < 0) id = authors.push(name) - 1;
		return id;
	};
	const sorted = [...comments].sort(
		(a, b) => a.address.row - b.address.row || a.address.col - b.address.col,
	);
	const threads = new ThreadedCommentsWriter(persons, threadOptions);
	const list = sorted
		.map((comment) => {
			const ref = formatAddress(comment.address);
			if (comment.replies) {
				const id = threads.add(comment, ref, sheetName);
				const fallback = `${THREADED_NOTE}${comment.text}${comment.replies.map((r) => `\nReply:\n    ${r.text}`).join('')}`;
				return `<comment ref="${ref}" authorId="${authorId(`tc=${id}`)}" shapeId="0" xr:uid="${id}"><text>${tElement(fallback)}</text></comment>`;
			}
			return `<comment ref="${ref}" authorId="${authorId(comment.author)}" shapeId="0"><text><r>${tElement(comment.text)}</r></text></comment>`;
		})
		.join('');
	const authorsXml = authors.map((name) => `<author>${escapeText(name)}</author>`).join('');
	const commentsXml = `${XML_HEADER}<comments xmlns="${NS.x}" xmlns:mc="${NS.mc}" mc:Ignorable="xr" xmlns:xr="${NS.xr}"><authors>${authorsXml}</authors><commentList>${list}</commentList></comments>`;
	const vml = vmlXml(
		sorted,
		layout ?? {
			blocks: [sheetNumber],
			ids: sorted.map((_, index) => 1024 * sheetNumber + index + 1),
		},
	);
	const parts: CommentParts = { comments: commentsXml, vml };
	const threaded = threads.xml();
	if (threaded) parts.threaded = threaded;
	return parts;
}
