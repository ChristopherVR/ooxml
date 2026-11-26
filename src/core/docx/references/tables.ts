import { isFieldLocked } from '../field-lock';
import { fieldName } from '../field-runs';
import type { Block, DocumentModel, Paragraph, TextRun } from '../model';
import { twipsFromPixels } from '../units';
import { authorityParagraphs } from './authorities';
import { bibliographyEntry } from './citations';
import { referenceEntries, type ReferencePage } from './entries';
import { indexParagraphs, INDEX_LEADERS, type IndexLeader } from './index-entries';
import { referenceFieldRuns, referenceInstruction } from './instructions';
import type { Bibliography } from './source-model';

export type ReferenceTable = 'INDEX' | 'TOA' | 'BIBLIOGRAPHY';

export interface ReferenceTableLocation {
	start: number;
	end: number;
	instruction: string;
	before: TextRun[];
	after: TextRun[];
	leader?: IndexLeader;
}

/** Like a TOC, these complex fields can begin and end in different body paragraphs. */
export function findReferenceTables(
	blocks: readonly Block[],
	kind: ReferenceTable,
): ReferenceTableLocation[] {
	const found: ReferenceTableLocation[] = [];
	const frames: { start: number; run: number; code: string; locked: boolean }[] = [];
	for (const [index, block] of blocks.entries()) {
		if (block.type !== 'paragraph') {
			frames.length = 0;
			continue;
		}
		for (const [at, run] of block.runs.entries()) {
			if (run.fieldChar === 'begin')
				frames.push({
					start: index,
					run: at,
					code: '',
					locked: isFieldLocked(run),
				});
			else if (run.fieldCode !== undefined) {
				const frame = frames.at(-1);
				if (frame && frame.code.length <= 8192) frame.code += run.fieldCode;
			} else if (run.fieldChar === 'end') {
				const frame = frames.pop();
				const first = frame && blocks[frame.start];
				if (!frame || frame.locked || first?.type !== 'paragraph' || fieldName(frame.code) !== kind)
					continue;
				const leader = first.tabStops?.find((tab) => tab.align === 'right')?.leader;
				found.push({
					start: frame.start,
					end: index,
					instruction: frame.code,
					before: first.runs.slice(0, frame.run),
					after: block.runs.slice(at + 1),
					...(INDEX_LEADERS.find((item) => item === leader)
						? { leader: INDEX_LEADERS.find((item) => item === leader)! }
						: {}),
				});
			}
		}
	}
	const ordered = found.sort((a, b) => a.start - b.start || b.end - a.end);
	let end = -1;
	return ordered.filter((item) => {
		if (item.start <= end) return false;
		end = item.end;
		return true;
	});
}

export function buildReferenceTable(
	model: DocumentModel,
	instruction: string,
	options: {
		page?: ReferencePage;
		bibliography?: Bibliography;
		leader?: IndexLeader;
		empty: string;
	},
): Paragraph[] {
	const code = referenceInstruction(instruction);
	const page = options.page ?? (() => undefined);
	const width = twipsFromPixels(model.page.width - model.page.marginLeft - model.page.marginRight);
	let paragraphs: Paragraph[] = [];
	if (code?.name === 'INDEX')
		paragraphs = indexParagraphs(
			referenceEntries(model),
			code,
			page,
			width,
			options.leader ?? 'dot',
		);
	else if (code?.name === 'TOA')
		paragraphs = authorityParagraphs(referenceEntries(model), code, page, width);
	else if (code?.name === 'BIBLIOGRAPHY' && options.bibliography) {
		const { style = 'APA', sources } = options.bibliography;
		const tags = new Set(code.values('m'));
		paragraphs = sources
			.filter((source) => !source.placeholder && (!tags.size || tags.has(source.tag)))
			.sort((a, b) =>
				(a.author?.kind === 'corporate'
					? a.author.name
					: a.author?.people[0]?.last || a.fields.Title || a.tag
				).localeCompare(
					b.author?.kind === 'corporate'
						? b.author.name
						: b.author?.people[0]?.last || b.fields.Title || b.tag,
				),
			)
			.map((source) => ({ type: 'paragraph', id: '', runs: bibliographyEntry(source, style) }));
	}
	if (!paragraphs.length)
		paragraphs.push({ type: 'paragraph', id: '', runs: [{ text: options.empty }] });
	const first = paragraphs[0];
	const last = paragraphs.at(-1);
	if (first && last) {
		const markers = referenceFieldRuns(instruction, []);
		for (const paragraph of paragraphs)
			paragraph.runs = paragraph.runs.map((run) => ({
				...run,
				field: { instr: instruction.trim() },
			}));
		first.runs.unshift(...markers.slice(0, -1));
		last.runs.push(...markers.slice(-1));
	}
	return paragraphs;
}
