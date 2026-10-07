import type { DocumentModel, Paragraph } from './model.js';
import {
	documentBlockLists,
	mapBlockParagraphs,
	mapDocumentParagraphs,
} from './document-paragraphs.js';

const WORD_MOVE_NAME = /^move([1-9]\d{0,8})$/;

/** Word's native reader links numeric move names; keep CRDT identities internal. */
export function numberMoveNames(model: DocumentModel): DocumentModel {
	const names = new Set<string>();
	for (const blocks of documentBlockLists(model))
		mapBlockParagraphs(blocks, (paragraph) => {
			for (const run of paragraph.runs) if (run.revision?.move) names.add(run.revision.move.name);
			return paragraph;
		});
	const invalid = [...names].filter((name) => !WORD_MOVE_NAME.test(name));
	if (!invalid.length) return model;
	let next =
		Math.max(0, ...[...names].map((name) => Number(WORD_MOVE_NAME.exec(name)?.[1] ?? 0))) + 1;
	const renamed = new Map(invalid.map((name) => [name, `move${next++}`]));
	const paragraph = (block: Paragraph): Paragraph => ({
		...block,
		runs: block.runs.map((run) => {
			const move = run.revision?.move;
			const name = move && renamed.get(move.name);
			return name ? { ...run, revision: { ...run.revision!, move: { ...move!, name } } } : run;
		}),
	});
	return mapDocumentParagraphs(model, paragraph);
}
