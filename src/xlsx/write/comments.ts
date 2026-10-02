import { NS } from '../../xml/index.js';
import { formatAddress } from '../address.js';
import type { Comment } from '../model.js';
import { XML_HEADER, escapeAttr, escapeText, tElement } from './xml-out.js';

/** A deterministic GUID-shaped id (`{XXXXXXXX-XXXX-...}`) from a seed string. */
export function guidFrom(seed: string): string {
	let h1 = 0x811c9dc5;
	let h2 = 0x01000193;
	const hex: string[] = [];
	for (let round = 0; round < 4; round++) {
		for (let i = 0; i < seed.length; i++) {
			h1 = Math.imul(h1 ^ seed.charCodeAt(i), 0x01000193) >>> 0;
			h2 = Math.imul(h2 ^ (seed.charCodeAt(i) + round), 0x5bd1e995) >>> 0;
		}
		h1 = Math.imul(h1 ^ round, 0x27d4eb2d) >>> 0;
		hex.push(((h1 ^ h2) >>> 0).toString(16).toUpperCase().padStart(8, '0'));
	}
	const s = hex.join('');
	return `{${s.slice(0, 8)}-${s.slice(8, 12)}-4${s.slice(13, 16)}-8${s.slice(17, 20)}-${s.slice(20, 32)}}`;
}

const THREADED_NOTE =
	'[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ';

export interface CommentParts {
	comments: string;
	vml: string;
	/** Present when at least one comment is threaded (has a `replies` array). */
	threaded?: string;
}

/** Persons referenced by threaded comments: display name to person id. */
export class PersonRegistry {
	private readonly names = new Map<string, string>();
	private readonly byName = new Map<string, string>();
	constructor(existing?: ReadonlyMap<string, string>) {
		for (const [id, name] of existing ?? []) {
			this.names.set(id, name);
			if (!this.byName.has(name)) this.byName.set(name, id);
		}
	}
	id(name: string): string {
		let id = this.byName.get(name);
		if (!id) {
			id = guidFrom(`person:${name}`);
			this.byName.set(name, id);
			this.names.set(id, name);
		}
		return id;
	}
	xml(): string {
		const people = [...this.names]
			.map(
				([id, name]) =>
					`<person displayName="${escapeAttr(name)}" id="${id}" userId="${escapeAttr(name)}" providerId="None"/>`,
			)
			.join('');
		return `${XML_HEADER}<personList xmlns="${NS.tc}" xmlns:x="${NS.x}">${people}</personList>`;
	}
}

function vmlShape(comment: Comment, index: number, idBase: number): string {
	const { row, col } = comment.address;
	const anchor = `${col + 1}, 15, ${Math.max(0, row - 1)}, 10, ${col + 3}, 15, ${row + 3}, 4`;
	return (
		`<v:shape id="_x0000_s${idBase + index + 1}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${index + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
		'<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
		'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
		`<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`
	);
}

/**
 * The comments part, the legacy VML drawing that makes Excel show them, and (for threaded
 * comments) the threaded comments part. `sheetNumber` keeps VML shape ids unique per sheet.
 */
export function commentParts(
	comments: readonly Comment[],
	sheetNumber: number,
	persons: PersonRegistry,
	sheetName: string,
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
	let threaded = '';
	const list = sorted
		.map((comment) => {
			const ref = formatAddress(comment.address);
			if (comment.replies) {
				const id = guidFrom(`${sheetName}!${ref}:${comment.text}`);
				const personId = persons.id(comment.author);
				threaded += `<threadedComment ref="${ref}" personId="${personId}" id="${id}"><text>${escapeText(comment.text)}</text></threadedComment>`;
				comment.replies.forEach((reply, index) => {
					const replyId = guidFrom(`${id}:reply:${index}:${reply.text}`);
					const dT = reply.date ? ` dT="${escapeAttr(reply.date)}"` : '';
					threaded += `<threadedComment ref="${ref}"${dT} personId="${persons.id(reply.author)}" id="${replyId}" parentId="${id}"><text>${escapeText(reply.text)}</text></threadedComment>`;
				});
				const fallback = `${THREADED_NOTE}${comment.text}${comment.replies.map((r) => `\nReply:\n    ${r.text}`).join('')}`;
				return `<comment ref="${ref}" authorId="${authorId(`tc=${id}`)}" shapeId="0" xr:uid="${id}"><text>${tElement(fallback)}</text></comment>`;
			}
			return `<comment ref="${ref}" authorId="${authorId(comment.author)}" shapeId="0"><text><r>${tElement(comment.text)}</r></text></comment>`;
		})
		.join('');
	const authorsXml = authors.map((name) => `<author>${escapeText(name)}</author>`).join('');
	const commentsXml = `${XML_HEADER}<comments xmlns="${NS.x}" xmlns:mc="${NS.mc}" mc:Ignorable="xr" xmlns:xr="${NS.xr}"><authors>${authorsXml}</authors><commentList>${list}</commentList></comments>`;
	const idBase = 1024 * sheetNumber;
	const vml =
		`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
		`<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheetNumber}"/></o:shapelayout>` +
		'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
		sorted.map((comment, index) => vmlShape(comment, index, idBase)).join('') +
		'</xml>';
	const parts: CommentParts = { comments: commentsXml, vml };
	if (threaded)
		parts.threaded = `${XML_HEADER}<ThreadedComments xmlns="${NS.tc}" xmlns:x="${NS.x}">${threaded}</ThreadedComments>`;
	return parts;
}
