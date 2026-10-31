// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Comment } from './model';
import { getW, parseXml, textContent, type XmlElement, WORD_NS, WORD_DATE_UTC_NS } from './xml';
import { W14_NS, W15_NS, W16CID_NS, W16CEX_NS } from './comment-parts';

function paragraphText(paragraph: XmlElement): string {
	return Array.from(paragraph.getElementsByTagNameNS(WORD_NS, 't'))
		.map((node) => textContent(node as XmlElement))
		.join('');
}

/** The modern metadata parts identify a comment by its last paragraph. */
export function commentParagraphId(element: XmlElement): string | undefined {
	const paragraphs = element.getElementsByTagNameNS(WORD_NS, 'p');
	return paragraphs[paragraphs.length - 1]?.getAttributeNS(W14_NS, 'paraId') || undefined;
}

/** Parses comments and their done/reply, persistent identity and UTC timestamp metadata. */
export function parseComments(
	commentsXml: string,
	extendedXml?: string,
	idsXml?: string,
	extensibleXml?: string,
): Comment[] {
	const doc = parseXml(commentsXml);
	const commentElements = Array.from(
		doc.getElementsByTagNameNS(WORD_NS, 'comment'),
	) as XmlElement[];
	const paraIdToCommentId = new Map<string, string>();
	const modernDates = new Set<string>();
	const comments: Comment[] = commentElements.map((element) => {
		const id = getW(element, 'id') ?? '0';
		const text = Array.from(element.getElementsByTagNameNS(WORD_NS, 'p'))
			.map((p) => paragraphText(p as XmlElement))
			.join('\n')
			.trim();
		const comment: Comment = { id, author: getW(element, 'author') || 'Unknown', text };
		const initials = getW(element, 'initials');
		if (initials) comment.initials = initials;
		const dateUtc = element.getAttributeNS(WORD_DATE_UTC_NS, 'dateUtc');
		if (dateUtc) modernDates.add(id);
		const date = dateUtc || getW(element, 'date');
		if (date) comment.date = date;
		const paraId = commentParagraphId(element);
		if (paraId) paraIdToCommentId.set(paraId.toUpperCase(), id);
		return comment;
	});
	const byId = new Map(comments.map((comment) => [comment.id, comment]));
	const byParagraph = (id: string | null) =>
		byId.get(paraIdToCommentId.get(id?.toUpperCase() ?? '') ?? '');
	if (extendedXml) {
		const extended = parseXml(extendedXml);
		for (const entry of Array.from(
			extended.getElementsByTagNameNS(W15_NS, 'commentEx'),
		) as XmlElement[]) {
			const paraId = entry.getAttributeNS(W15_NS, 'paraId');
			const comment = byParagraph(paraId);
			if (!comment) continue;
			const done = entry.getAttributeNS(W15_NS, 'done');
			comment.resolved = done === '1' || done === 'true' || done === 'on';
			const parentParaId = entry.getAttributeNS(W15_NS, 'paraIdParent');
			const parentId = byParagraph(parentParaId)?.id;
			if (parentId) comment.parentId = parentId;
		}
	}
	if (idsXml && extensibleXml) {
		const byDurableId = new Map<string, Comment>();
		for (const entry of Array.from(
			parseXml(idsXml).getElementsByTagNameNS(W16CID_NS, 'commentId'),
		)) {
			const comment = byParagraph(entry.getAttributeNS(W16CID_NS, 'paraId'));
			const durableId = entry.getAttributeNS(W16CID_NS, 'durableId');
			if (comment && durableId) byDurableId.set(durableId.toUpperCase(), comment);
		}
		for (const entry of Array.from(
			parseXml(extensibleXml).getElementsByTagNameNS(W16CEX_NS, 'commentExtensible'),
		)) {
			const comment = byDurableId.get(
				entry.getAttributeNS(W16CEX_NS, 'durableId')?.toUpperCase() ?? '',
			);
			const date = entry.getAttributeNS(W16CEX_NS, 'dateUtc');
			if (comment && date && !modernDates.has(comment.id)) comment.date = date;
		}
	}
	return comments;
}
