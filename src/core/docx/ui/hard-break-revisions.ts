import type { Node } from 'prosemirror-model';
import type { Transform } from 'prosemirror-transform';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';

/** Text marks on line-break elements must be stored as run attributes before Yjs sees them. */
export function canonicalHardBreak(node: Node): Node {
	return node.type.name === 'hardBreak' && node.type.spec.attrs?.format
		? runToInlineNodes(inlineNodeRun(node)!, node.type.schema)[0]!
		: node;
}

export function canonicalizeHardBreaks(tr: Transform, from: number, to: number): void {
	const breaks: { pos: number; node: Node }[] = [];
	tr.doc.nodesBetween(from, to, (node, pos) => {
		const canonical = canonicalHardBreak(node);
		if (!node.eq(canonical)) breaks.push({ pos, node: canonical });
	});
	for (const { pos, node } of breaks) tr.setNodeMarkup(pos, undefined, node.attrs, node.marks);
}
