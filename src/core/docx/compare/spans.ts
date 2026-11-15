import type { Block, DocumentModel } from '../model';

/** One tracked span of a compared document: its kind and its text (`¶` for a paragraph mark). */
export interface RevisionSpan {
	kind: 'insert' | 'delete' | 'formatChange' | 'unchanged';
	text: string;
}

/**
 * The body as an ordered list of spans, adjacent spans of one kind merged: paragraph marks are
 * `¶`, a row revision `[row]`. Two comparisons that agree on what was inserted and deleted have
 * the same spans, whatever their run boundaries or revision ids.
 */
export function revisionSpans(model: DocumentModel): RevisionSpan[] {
	const spans: RevisionSpan[] = [];
	const push = (kind: RevisionSpan['kind'], text: string) => {
		const last = spans.at(-1);
		if (last?.kind === kind) last.text += text;
		else spans.push({ kind, text });
	};
	const kindOf = (kind: string | undefined): RevisionSpan['kind'] =>
		kind === 'insert' || kind === 'moveTo'
			? 'insert'
			: kind === 'delete' || kind === 'moveFrom'
				? 'delete'
				: kind === 'formatChange'
					? 'formatChange'
					: 'unchanged';
	const visit = (blocks: Block[]) => {
		for (const block of blocks) {
			if (block.type === 'table') {
				block.rows.forEach((row, index) => {
					const revision = block.rowProperties?.[index]?.revision;
					if (revision) push(kindOf(revision.kind), '[row]');
					for (const cell of row) visit(cell.paragraphs);
				});
				continue;
			}
			for (const run of block.runs)
				if (run.text) push(kindOf(run.formatRevision?.kind ?? run.revision?.kind), run.text);
			push(kindOf(block.markRevision?.kind), '¶');
		}
	};
	visit(model.blocks);
	return spans;
}
