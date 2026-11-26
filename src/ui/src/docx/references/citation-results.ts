import { Fragment, type Node as ProseMirrorNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { DocumentModel } from 'ooxml-core/docx';
import { runToInlineNodes, type Bibliography, type WordReferences } from 'ooxml-core/docx/ui';
import { schema } from '../schema';

export function refreshCitationResults(
	references: WordReferences,
	tr: Transaction,
	model: DocumentModel,
	bibliography: Bibliography,
) {
	const evaluate = references.citationEvaluator(bibliography);
	const paragraphs = new Map<
		string,
		Extract<DocumentModel['blocks'][number], { type: 'paragraph' }>
	>();
	for (const block of model.blocks) {
		if (block.type === 'paragraph') paragraphs.set(block.id, block);
		else
			for (const row of block.rows)
				for (const cell of row)
					for (const paragraph of cell.paragraphs) paragraphs.set(paragraph.id, paragraph);
	}
	const changes: { pos: number; node: ProseMirrorNode; next: Fragment }[] = [];
	tr.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		const paragraph = paragraphs.get(String(node.attrs.id));
		if (!paragraph) return false;
		const updated = references.refreshCitations(paragraph, evaluate);
		if (updated !== paragraph)
			changes.push({
				pos,
				node,
				next: Fragment.fromArray(updated.runs.flatMap((run) => runToInlineNodes(run, schema))),
			});
		return false;
	});
	for (const { pos, node, next } of changes.reverse()) {
		const first = node.content.findDiffStart(next);
		const last = node.content.findDiffEnd(next);
		if (first !== null && last) {
			const overlap = Math.max(0, first - Math.min(last.a, last.b));
			tr.replaceWith(
				pos + 1 + first,
				pos + 1 + last.a + overlap,
				next.cut(first, last.b + overlap),
			);
		}
	}
}
