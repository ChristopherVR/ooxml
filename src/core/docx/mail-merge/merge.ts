import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	HeaderFooterSlots,
	Note,
	SectionProperties,
	TextRun,
} from '../model';
import { rememberBlockCopies, type BlockCopy } from '../block-copies';
import { documentParagraphs, mapStoryBlocks } from '../cell-blocks';
import { sectionsOf } from '../section-layout';
import { selectMailMergeRecipients, type MailMergeData } from './data-source';
import type { MailMergeIssue, MailMergeOptions } from './field-values';
import { projectMerge } from './project';

interface MergeResult {
	issues: MailMergeIssue[];
}

export interface MailMergePreview extends MergeResult {
	/** The record number used to start this output, one-based in the selected list. */
	recordNumber: number;
	document: DocumentModel;
	skipped: boolean;
}

export interface MailMergeIndividualResult extends MergeResult {
	documents: { recordNumber: number; document: DocumentModel }[];
}

export interface MailMergeCombinedResult extends MergeResult {
	/** Absent if the list is empty or SKIPIF discarded every output. */
	document?: DocumentModel;
	mergedCount: number;
}

function optionsFor(model: DocumentModel, options: MailMergeOptions): MailMergeOptions {
	const fieldMappings = options.fieldMappings ?? model.mailMerge?.fieldMappings;
	return fieldMappings ? { ...options, fieldMappings } : options;
}

/** Projects a one-based recipient onto a view model. NEXT can also consume later recipients. */
export function previewMailMerge(
	model: DocumentModel,
	data: MailMergeData,
	recordNumber: number,
	options: MailMergeOptions = {},
): MailMergePreview {
	if (!Number.isSafeInteger(recordNumber) || recordNumber < 1 || recordNumber > data.records.length)
		throw new RangeError(`No mail merge record at position ${recordNumber}.`);
	const prepared = selectMailMergeRecipients(data);
	let sequence = 1;
	let index = 0;
	// Count completed documents before this preview, including skipped and consumed records.
	while (index < recordNumber - 1) {
		const projected = projectMerge(model, prepared, index, sequence, optionsFor(model, options));
		if (projected.nextIndex >= recordNumber) break;
		if (!projected.skipped) sequence++;
		index = projected.nextIndex;
	}
	const projected = projectMerge(
		model,
		prepared,
		recordNumber - 1,
		sequence,
		optionsFor(model, options),
	);
	return {
		recordNumber,
		document: projected.model,
		skipped: projected.skipped,
		issues: projected.issues,
	};
}

function* projectRecipients(model: DocumentModel, data: MailMergeData, options: MailMergeOptions) {
	const prepared = selectMailMergeRecipients(data);
	const resolved = optionsFor(model, options);
	let sequence = 1;
	for (let index = 0; index < prepared.records.length;) {
		const project = projectMerge(model, prepared, index, sequence, resolved);
		yield { ...project, recordNumber: index + 1 };
		if (!project.skipped) sequence++;
		index = project.nextIndex;
	}
}

function initialIssues(data: MailMergeData): MailMergeIssue[] {
	return data.records.length ? [] : [{ code: 'no-recipients' }];
}

/** Dry-runs the same field and record rules that finish uses. */
export function checkMailMerge(
	model: DocumentModel,
	data: MailMergeData,
	options: MailMergeOptions = {},
): MailMergeIssue[] {
	const issues = initialIssues(data);
	for (const project of projectRecipients(model, data, options)) issues.push(...project.issues);
	return issues;
}

function detachSettings({ mailMerge: _mailMerge, ...document }: DocumentModel): DocumentModel {
	return document;
}

/** Produces an independent document model for each output of the merge. */
export function finishMailMergeIndividually(
	model: DocumentModel,
	data: MailMergeData,
	options: MailMergeOptions = {},
): MailMergeIndividualResult {
	const documents: MailMergeIndividualResult['documents'] = [];
	const issues = initialIssues(data);
	for (const project of projectRecipients(model, data, options)) {
		issues.push(...project.issues);
		if (!project.skipped)
			documents.push({
				recordNumber: project.recordNumber,
				document: detachSettings(project.model),
			});
	}
	return { issues, documents };
}

function noteIds(notes: readonly Note[] | undefined) {
	const used = new Set(notes?.map((note) => note.id));
	let id = 1;
	return () => {
		while (used.has(String(id))) id++;
		return String(id++);
	};
}

function copySlots(
	slots: HeaderFooterSlots | undefined,
	inherited: HeaderFooterSlots | undefined,
	partName: (source: string, slot: string) => string,
): HeaderFooterSlots | undefined {
	if (!slots && !inherited) return undefined;
	const copy: HeaderFooterSlots = {};
	for (const slot of ['default', 'first', 'even'] as const) {
		const content: HeaderFooterContent | undefined =
			slots?.[slot] ??
			(inherited?.[slot] && {
				blocks: [{ type: 'paragraph' as const, id: 'p0', runs: [{ text: '' }] }],
			});
		if (!content) continue;
		const source = content.partName;
		copy[slot] = {
			...content,
			partName: partName(source ?? '', slot),
			...(source && { sourcePartName: content.sourcePartName ?? source }),
		};
	}
	return Object.keys(copy).length ? copy : undefined;
}

/** Concatenates merged copies, starting each on a new page and retaining internal section breaks. */
export function finishMailMerge(
	model: DocumentModel,
	data: MailMergeData,
	options: MailMergeOptions = {},
): MailMergeCombinedResult {
	const issues = initialIssues(data);
	const blocks: Block[] = [];
	const copies = new Map<string, BlockCopy>();
	const sourceIds = [...model.blocks, ...documentParagraphs(model.blocks)].map((block) => block.id);
	let prefix = 'mail-merge-';
	while (sourceIds.some((id) => id.startsWith(prefix))) prefix += '-';
	const sections: SectionProperties[] = [];
	const footnotes: Note[] = [];
	const endnotes: Note[] = [];
	const nextFootnote = noteIds(model.footnotes);
	const nextEndnote = noteIds(model.endnotes);
	const partNames = new Map<string, string>();
	const templateSections = sectionsOf(model);
	const lastHeaders: HeaderFooterSlots = {};
	const lastFooters: HeaderFooterSlots = {};
	for (const section of templateSections) {
		Object.assign(lastHeaders, section.headers);
		Object.assign(lastFooters, section.footers);
	}
	const usedParts = new Set(
		templateSections.flatMap((section) =>
			[...Object.values(section.headers ?? {}), ...Object.values(section.footers ?? {})].map(
				(part) => part?.partName,
			),
		),
	);
	let mergedCount = 0;
	for (const project of projectRecipients(model, data, options)) {
		issues.push(...project.issues);
		if (project.skipped) continue;
		const copy = project.model;
		const first = mergedCount === 0;
		const previous = blocks.at(-1);
		if (previous?.type === 'table') {
			const id = `${prefix}break-${mergedCount}`;
			blocks.push({ type: 'paragraph', id, runs: [{ text: '' }] });
			const section = sections.at(-1);
			if (section) section.endsAtBlockId = id;
		}
		const recordPrefix = `${prefix}${mergedCount}-`;
		const rekey = (id: string) => (first ? id : `${recordPrefix}${id}`);
		if (!first) for (const { id } of copy.blocks) copies.set(rekey(id), { sourceId: id, rekey });
		const copyNotes = (source: Note[] | undefined, into: Note[], allocate: () => string) => {
			const ids = new Map<string, string>();
			for (const note of source ?? []) {
				const id = first ? note.id : allocate();
				ids.set(note.id, id);
				into.push({ ...note, id });
			}
			return ids;
		};
		const noteRefs = {
			footnote: copyNotes(copy.footnotes, footnotes, nextFootnote),
			endnote: copyNotes(copy.endnotes, endnotes, nextEndnote),
		};
		const renameNote = (run: TextRun): TextRun => {
			const ref = run.noteReference;
			const id = ref && noteRefs[ref.kind].get(ref.id);
			return id && id !== ref.id ? { ...run, noteReference: { ...ref, id } } : run;
		};
		blocks.push(
			...(first
				? copy.blocks
				: mapStoryBlocks(
						copy.blocks,
						(paragraph) => ({
							...paragraph,
							id: rekey(paragraph.id),
							runs: paragraph.runs.map(renameNote),
						}),
						rekey,
					)),
		);
		sectionsOf(copy).forEach((section, index, all) => {
			const partName = (kind: string, source: string, slot: string) => {
				const key = `${mergedCount}:${kind}:${source || `${index}:${slot}`}`;
				const known = partNames.get(key);
				if (known) return known;
				let name = `word/mail-merge-${kind}-${partNames.size + 1}.xml`;
				while (usedParts.has(name)) name = name.replace('.xml', '-copy.xml');
				usedParts.add(name);
				partNames.set(key, name);
				return name;
			};
			const { headers, footers, ...layout } = section;
			// A fresh recipient's first section must not inherit a previous recipient's final story.
			const startsCopy = !first && index === 0;
			const copiedHeaders = copySlots(headers, startsCopy ? lastHeaders : undefined, (part, slot) =>
				partName('header', part, slot),
			);
			const copiedFooters = copySlots(footers, startsCopy ? lastFooters : undefined, (part, slot) =>
				partName('footer', part, slot),
			);
			sections.push({
				...layout,
				endsAtBlockId:
					index === all.length - 1 ? (blocks.at(-1)?.id ?? '') : rekey(section.endsAtBlockId),
				type: startsCopy ? 'nextPage' : layout.type,
				...(copiedHeaders && { headers: copiedHeaders }),
				...(copiedFooters && { footers: copiedFooters }),
			});
		});
		mergedCount++;
	}
	if (!mergedCount) return { issues, mergedCount };
	const document = {
		...detachSettings(model),
		blocks,
		sections,
		...(model.footnotes && { footnotes }),
		...(model.endnotes && { endnotes }),
	};
	rememberBlockCopies(document, copies);
	return { issues, mergedCount, document };
}
