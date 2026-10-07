import type { TextRun } from './model';
import { rangeStartNodes, rangeEndNodes, type RangeEdge } from './write-ranges';
import type { XmlDocument, XmlElement } from './xml';

/** Whole-field comments enclose the wrapper so Word includes the field code in their scope. */
export function takeSimpleFieldCommentEdges(
	doc: XmlDocument,
	runs: TextRun[],
	start: number,
	end: number,
	opens: Map<number, RangeEdge[]>,
	closes: Map<number, RangeEdge[]>,
): { before: XmlElement[]; after: XmlElement[] } {
	const coversField = (edge: RangeEdge) =>
		edge.kind === 'comment' &&
		runs.slice(start, end).every((run) => run.commentIds?.includes(edge.id));
	const take = (map: Map<number, RangeEdge[]>, index: number) => {
		const edges = map.get(index) ?? [];
		map.set(
			index,
			edges.filter((edge) => !coversField(edge)),
		);
		return edges.filter(coversField);
	};
	return {
		before: take(opens, start).flatMap((edge) => rangeStartNodes(doc, edge)),
		after: take(closes, end - 1)
			.reverse()
			.flatMap((edge) => rangeEndNodes(doc, edge)),
	};
}
