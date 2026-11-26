import type { Block, DocumentModel } from '../model';
import { cellSourceBlocks } from '../cell-blocks';
import type { MailMergeData } from './data-source';

/** Bounds a UI merge before projecting records or allocating repeated template blocks. */
export function mailMergeWithinWorkLimit(model: DocumentModel, data: MailMergeData): boolean {
	const WORK_LIMIT = 2_000_000;
	let work = 0;
	let fields = 0;
	let nativeFields = 0;
	let promptResults = 0;
	const visit = (blocks: readonly Block[], depth = 0): boolean => {
		if (depth > 32) return false;
		for (const block of blocks) {
			work += 32;
			if (block.type === 'paragraph') {
				for (const run of block.runs) {
					work += 16 + run.text.length + (run.fieldCode?.length ?? 0) + (run.image ? 4096 : 0);
					if (run.field?.simple || run.fieldChar === 'begin') fields++;
					const instruction = run.field?.simple ? run.field.instr : run.fieldCode;
					if (instruction) {
						if (/^\s*(?:ADDRESSBLOCK|GREETINGLINE)\b/i.test(instruction)) nativeFields++;
						if (/^\s*(?:FILLIN|REF)\b/i.test(instruction)) promptResults++;
					}
					if (work * data.records.length > WORK_LIMIT) return false;
				}
			} else {
				for (const row of block.rows)
					for (const cell of row) if (!visit(cellSourceBlocks(cell), depth + 1)) return false;
			}
			if (work * data.records.length > WORK_LIMIT) return false;
		}
		return true;
	};
	if (!data.records.length) return true;
	if (!visit(model.blocks)) return false;
	for (const section of model.sections ?? [])
		for (const story of [
			...Object.values(section.headers ?? {}),
			...Object.values(section.footers ?? {}),
		])
			if (story && !visit(story.blocks)) return false;
	for (const note of [...(model.footnotes ?? []), ...(model.endnotes ?? [])])
		if (!visit(note.blocks)) return false;
	// A field may repeat a record's longest value. Count those characters before allocating copies.
	let output = (work + promptResults * 1024) * data.records.length;
	for (const record of data.records) {
		let longest = 0;
		for (const value of Object.values(record))
			if (typeof value === 'string') longest = Math.max(longest, value.length);
		output += longest * fields + Math.min(longest, 1024) * nativeFields * 64;
		if (output > WORK_LIMIT) return false;
	}
	return true;
}
