import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { Revision, TextRun } from '../model';
import { restoreRunFormatting } from '../restore-run-format';
import { marksForRun } from './run-marks';
import { inlineNodeRun } from './run-adapter';
import { RUN_FORMAT_MARKS as FORMAT_MARKS, setInlineRunFormatting } from './inline-formatting';

/** Imported formatting history travels with run marks or inline object attributes. */
export function formattingRevision(node: ProseMirrorNode): Revision | undefined {
	if (node.type.name === 'equation') return undefined;
	const props = inlineNodeRun(node);
	const revision = props?.formatRevision ?? props?.revision;
	return revision?.kind === 'formatChange' ? revision : undefined;
}

/** Resolves formatting only, retaining text, links, comments and other inline marks. */
export function resolveFormattingRange(
	tr: Transaction,
	from: number,
	to: number,
	mode: 'accept' | 'reject',
): void {
	const pieces: { from: number; to: number; node: ProseMirrorNode; revision: Revision }[] = [];
	tr.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isInline) return;
		const revision = formattingRevision(node);
		if (revision)
			pieces.push({
				from: Math.max(from, pos),
				to: Math.min(to, pos + node.nodeSize),
				node,
				revision,
			});
	});
	for (const piece of pieces) {
		if (!piece.node.isText && piece.node.type.name !== 'hardBreak') {
			const run = inlineNodeRun(piece.node)!;
			if (mode === 'reject') restoreRunFormatting(run);
			else if (run.formatRevision) delete run.formatRevision;
			else delete run.revision;
			setInlineRunFormatting(tr, piece.from, run);
			continue;
		}
		const properties = piece.node.marks.find((mark) => mark.type.name === 'runProperties')!;
		if (mode === 'accept') {
			const props = { ...properties.attrs.props } as TextRun;
			if (props.formatRevision) delete props.formatRevision;
			else delete props.revision;
			tr.removeMark(piece.from, piece.to, properties);
			if (Object.keys(props).length)
				tr.addMark(piece.from, piece.to, properties.type.create({ props }));
			continue;
		}
		const run: TextRun = { text: '', revision: piece.revision };
		restoreRunFormatting(run);
		for (const mark of piece.node.marks)
			if (FORMAT_MARKS.has(mark.type.name)) tr.removeMark(piece.from, piece.to, mark);
		for (const mark of marksForRun(run, tr.doc.type.schema)) tr.addMark(piece.from, piece.to, mark);
	}
}
