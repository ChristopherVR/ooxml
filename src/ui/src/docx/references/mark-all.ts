import type { Node } from 'prosemirror-model';
import { inlineNodeRun } from 'ooxml-core/docx/ui';

/** One case-sensitive, whole-word match per paragraph, outside any field. */
export function markAllPositions(doc: Node, selected: string): number[] {
	if (!selected) return [];
	const pattern = selected.replace(/[.*+?^$(){}|[\]\\]/g, '\\$&');
	const boundary = '[\\p{L}\\p{M}\\p{N}_]';
	const word = new RegExp('(?<!' + boundary + ')' + pattern + '(?!' + boundary + ')', 'u');
	const positions: number[] = [];
	let fieldDepth = 0;
	doc.descendants((paragraph, pos) => {
		if (positions.length >= 10_000) return false;
		if (paragraph.type.name !== 'paragraph') return true;
		const chunks: string[] = [];
		paragraph.forEach((node) => {
			const run = inlineNodeRun(node);
			if (run?.fieldChar === 'begin') fieldDepth++;
			else if (run?.fieldChar === 'end') fieldDepth = Math.max(0, fieldDepth - 1);
			const text = !fieldDepth && !run?.field && node.text;
			chunks.push(text || '\0'.repeat(node.nodeSize));
		});
		const match = word.exec(chunks.join(''));
		if (match) positions.push(pos + 1 + match.index + match[0].length);
		return false;
	});
	return positions;
}
