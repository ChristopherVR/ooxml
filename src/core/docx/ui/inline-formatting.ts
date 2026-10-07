import type { Transaction } from 'prosemirror-state';
import type { TextRun } from '../model';
import { runToInlineNodes } from './run-adapter';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { updatedInlineRunAttributes } from './inline-run-properties';

export const hasInlineRunAttributes = (node: ProseMirrorNode): boolean =>
	!node.isText && Boolean(node.type.spec.attrs?.format);

export const RUN_FORMAT_MARKS = new Set([
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

/** Atom properties live in attributes because the Yjs mapping does not retain element marks. */
export function setInlineRunFormatting(tr: Transaction, pos: number, run: TextRun): void {
	const node = tr.doc.nodeAt(pos)!;
	const projected = runToInlineNodes(run, tr.doc.type.schema)[0]!;
	const attrs = updatedInlineRunAttributes(node, projected.attrs.format);
	for (const [key, value] of Object.entries(attrs))
		if (node.attrs[key] !== value) tr.setNodeAttribute(pos, key, value);
	for (const mark of node.marks)
		if (RUN_FORMAT_MARKS.has(mark.type.name)) tr.removeMark(pos, pos + node.nodeSize, mark.type);
}
