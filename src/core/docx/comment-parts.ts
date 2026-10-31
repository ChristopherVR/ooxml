import type JSZip from 'jszip';
import { parseRelationships, resolvePartPath } from '../opc/relationships';

export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
export const W16CID_NS = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
export const W16CEX_NS = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';

export const COMMENT_PARTS = {
	comments: {
		name: 'comments',
		relationship: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
	},
	extended: {
		name: 'commentsExtended',
		relationship: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
	},
	ids: {
		name: 'commentsIds',
		relationship: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
	},
	extensible: {
		name: 'commentsExtensible',
		relationship: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
	},
} as const;
export type CommentPartKind = keyof typeof COMMENT_PARTS;
export interface CommentPart {
	path: string;
	xml: string | undefined;
	related: boolean;
}
export type CommentParts = Record<CommentPartKind, CommentPart>;

/** Follow source relationships, retaining the conventional-path fallback for minimal packages. */
export async function readCommentParts(zip: JSZip): Promise<CommentParts> {
	const relationships = parseRelationships(
		await zip.file('word/_rels/document.xml.rels')?.async('string'),
	);
	const entries = await Promise.all(
		Object.entries(COMMENT_PARTS).map(async ([kind, definition]) => {
			const relationship = [...relationships.values()].find(
				(rel) =>
					rel.mode === 'Internal' &&
					(rel.type === definition.relationship ||
						(kind === 'comments' &&
							rel.type === 'http://purl.oclc.org/ooxml/officeDocument/relationships/comments')),
			);
			const path = relationship
				? resolvePartPath('word/document.xml', relationship.target)
				: `word/${definition.name}.xml`;
			return [
				kind,
				{ path, xml: await zip.file(path)?.async('string'), related: !!relationship },
			] as const;
		}),
	);
	return Object.fromEntries(entries) as CommentParts;
}
