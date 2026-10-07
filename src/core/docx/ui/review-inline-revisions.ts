import type { Node } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { Revision, TextRun } from '../model';
import { inlineNodeRun } from './run-adapter';

export function inlineTextRevision(node: Node): Revision | undefined {
	const revision = inlineNodeRun(node)?.revision;
	return revision && ['insert', 'delete', 'moveTo', 'moveFrom'].includes(revision.kind)
		? revision
		: undefined;
}

/** Retained atoms carry history in their own properties instead of inheritable text marks. */
export function clearInlineTextRevisions(tr: Transaction, from: number, to: number): void {
	const atoms: { pos: number; format: Partial<TextRun> }[] = [];
	tr.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isInline || node.isText || typeof node.attrs.format !== 'string') return;
		const format = { ...JSON.parse(node.attrs.format) } as Partial<TextRun>;
		if (
			!format.revision ||
			!['insert', 'delete', 'moveTo', 'moveFrom'].includes(format.revision.kind)
		)
			return;
		delete format.revision;
		atoms.push({ pos, format });
	});
	for (const { pos, format } of atoms)
		tr.setNodeAttribute(pos, 'format', Object.keys(format).length ? JSON.stringify(format) : null);
}
