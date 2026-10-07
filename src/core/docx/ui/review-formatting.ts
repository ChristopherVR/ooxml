import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { Revision, TextRun } from '../model.js';
import { restoreRunFormatting } from '../restore-run-format.js';
import { marksForRun } from './run-marks.js';

/** Imported formatting history travels with the opaque run properties mark. */
export function formattingRevision(node: ProseMirrorNode): Revision | undefined {
	const revision = node.marks.find((mark) => mark.type.name === 'runProperties')?.attrs.props
		?.revision as Revision | undefined;
	return revision?.kind === 'formatChange' ? revision : undefined;
}

const FORMAT_MARKS = new Set([
	'bold',
	'italic',
	'underline',
	'strike',
	'highlight',
	'verticalAlign',
	'language',
	'runRtl',
	'font',
	'characterStyle',
	'runProperties',
]);

/** Resolves formatting only, retaining text, links, comments and other inline marks. */
export function resolveFormattingRange(
	tr: Transaction,
	from: number,
	to: number,
	mode: 'accept' | 'reject',
): void {
	const pieces: { from: number; to: number; node: ProseMirrorNode; revision: Revision }[] = [];
	tr.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return;
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
		const properties = piece.node.marks.find((mark) => mark.type.name === 'runProperties')!;
		if (mode === 'accept') {
			const { revision: _revision, ...props } = properties.attrs.props as TextRun;
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
