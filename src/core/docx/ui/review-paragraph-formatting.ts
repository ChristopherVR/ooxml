import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { Paragraph, Revision } from '../model';
import { PARAGRAPH_FORMAT_KEYS, restoreParagraphFormatting } from '../restore-paragraph-format';

export function paragraphFormattingRevision(node: ProseMirrorNode): Revision | undefined {
	const revision = node.attrs.formatRevision as Revision | undefined;
	return node.type.name === 'paragraph' && revision?.kind === 'paragraphChange'
		? revision
		: undefined;
}

/** Update paragraph attributes without replacing its text or changing its identity. */
export function resolveParagraphFormatting(
	tr: Transaction,
	pos: number,
	mode: 'accept' | 'reject',
): void {
	const node = tr.doc.nodeAt(pos);
	if (!node) return;
	const revision = paragraphFormattingRevision(node);
	if (!revision) return;
	const attrs: Record<string, unknown> = { ...node.attrs, formatRevision: null };
	if (mode === 'reject') {
		const paragraph: Paragraph = {
			type: 'paragraph',
			id: String(node.attrs.id),
			runs: [],
			formatRevision: revision,
			...(node.attrs.sourceParagraphPropertiesXml
				? { sourceParagraphPropertiesXml: String(node.attrs.sourceParagraphPropertiesXml) }
				: {}),
		};
		restoreParagraphFormatting(paragraph);
		for (const key of PARAGRAPH_FORMAT_KEYS)
			if (Object.hasOwn(node.attrs, key))
				attrs[key] = paragraph[key] ?? node.type.spec.attrs?.[key]?.default ?? null;
		attrs.numId = paragraph.numbering?.numId ?? null;
		attrs.ilvl = paragraph.numbering?.level ?? null;
		attrs.restoredParagraphPropertiesXml = paragraph.restoredParagraphPropertiesXml;
		if (paragraph.sourceParagraphPropertiesXml)
			attrs.sourceParagraphPropertiesXml = paragraph.sourceParagraphPropertiesXml;
	}
	tr.setNodeMarkup(pos, undefined, attrs);
}
