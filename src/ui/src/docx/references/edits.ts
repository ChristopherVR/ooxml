import { Fragment, type Node as ProseMirrorNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { sectionLayoutJson, type DocumentModel, type TextRun } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import {
	inlineNodeRun,
	runToInlineNodes,
	sectionPartsJson,
	type Bibliography,
	type IndexLeader,
	type ReferencePage,
	type ReferenceTable,
	type WordReferences,
} from 'ooxml-core/docx/ui';
import { createCanvasMeasurer } from '../canvas-measurer';
import { documentInformationJson } from '../document-information-attr';
import { modelToDoc } from '../model-adapter';
import { pageNumbers } from '../print-header-footer';
import { schema } from '../schema';

export function dispatchReference(view: EditorView, tr: Transaction): void {
	if (view.editable && tr.docChanged) view.dispatch(closeHistory(tr).scrollIntoView());
}

export function inlineReference(tr: Transaction, from: number, to: number, runs: TextRun[]): void {
	tr.replaceWith(
		from,
		to,
		runs.flatMap((run) => runToInlineNodes(run, schema)),
	);
}

export function markEntry(
	references: WordReferences,
	view: EditorView,
	code: string,
	all: string | undefined,
	format?: Pick<TextRun, 'bold' | 'italic' | 'underline'>,
): void {
	const tr = view.state.tr;
	const runs = references
		.referenceFieldRuns(code)
		.map((run) => (run.fieldCode ? { ...run, ...format } : run));
	if (!all) inlineReference(tr, tr.selection.to, tr.selection.to, runs);
	else {
		const pattern = all.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		const word = new RegExp(`(?<![\\p{L}\\p{M}\\p{N}_])${pattern}(?![\\p{L}\\p{M}\\p{N}_])`, 'giu');
		const positions: number[] = [];
		tr.doc.descendants((node, pos) => {
			if (positions.length >= 10_000) return false;
			if (node.type.name !== 'paragraph') return true;
			const text = node.textBetween(0, node.content.size, '', '\0');
			for (const match of text.matchAll(word)) {
				positions.push(pos + 1 + match.index + match[0].length);
				if (positions.length >= 10_000) break;
			}
			return false;
		});
		for (const at of positions.reverse()) inlineReference(tr, at, at, runs);
	}
	dispatchReference(view, tr);
}

export function bibliographyOf(
	references: WordReferences,
	model: DocumentModel,
): Bibliography | undefined {
	const part = model.packageInformation?.bibliography;
	if (!part) return { sources: [], style: 'APA' };
	const data = references.readBibliographyXml(part.xml);
	return data && { ...data, ...(part.part ? { part: part.part } : {}) };
}

export function setBibliography(
	references: WordReferences,
	tr: Transaction,
	model: DocumentModel,
	data: Bibliography,
): DocumentModel {
	const original = model.packageInformation?.bibliography;
	const next = {
		...model,
		packageInformation: {
			...model.packageInformation,
			bibliography: { ...original, xml: references.writeBibliographyXml(data, original?.xml) },
		},
	};
	tr.setDocAttribute('documentInformation', documentInformationJson(next));
	return next;
}

type PagePart = { number: string; end: number };

function referencePages(model: DocumentModel): ReferencePage {
	const { pages } = layoutDocumentModel(model, createCanvasMeasurer());
	const numbers = pageNumbers(model, pages);
	const fragments = new Map<string, PagePart[]>();
	const add = (id: string, number: string, end: number) => {
		const parts = fragments.get(id) ?? [];
		parts.push({ number, end });
		fragments.set(id, parts);
	};
	for (const [index, page] of pages.entries()) {
		const number = numbers[index];
		if (number === undefined) continue;
		for (const column of page.columns)
			for (const block of column.blocks) {
				if (block.kind === 'paragraph')
					add(block.blockId, number, block.lines.at(-1)?.sourceEnd ?? Infinity);
				else
					for (const row of block.rows) {
						if (row.repeated) continue;
						for (const cell of row.cells)
							for (const box of cell)
								add(box.blockId, number, box.lines.at(-1)?.sourceEnd ?? Infinity);
					}
			}
	}
	return (id, offset, end) => {
		const parts = fragments.get(id);
		if (end || offset === undefined) return (end ? parts?.at(-1) : parts?.[0])?.number;
		return (parts?.find((part) => offset < part.end) ?? parts?.at(-1))?.number;
	};
}

/** Rebuild each result on demand using the same paginated model as the table of contents. */
export function changeReferenceTables(
	references: WordReferences,
	tr: Transaction,
	model: DocumentModel,
	options: {
		kind: ReferenceTable;
		instruction?: string;
		leader?: IndexLeader;
		columns?: number;
		bibliography?: Bibliography;
		heading?: string;
		empty: string;
	},
): boolean {
	const { kind, instruction, empty, bibliography } = options;
	const { $from } = tr.selection;
	const start = $from.index(0);
	const top = tr.doc.maybeChild(start);
	const edit = instruction
		? [
				{
					start,
					end: top?.type.name === 'paragraph' && top.content.size === 0 ? start : start - 1,
					instruction,
					before: [],
					after: [],
					leader: options.leader,
				},
			]
		: references.findReferenceTables(model.blocks, kind);
	if (!edit.length) return false;
	const newId = () => `ref-${references.createOfficeGuid()}`;
	const ordered = [...edit].sort((a, b) => b.start - a.start);
	const build = (page?: ReferencePage) =>
		ordered.map((item) => {
			const paragraphs = references.buildReferenceTable(model, item.instruction, {
				empty,
				blockIndex: item.start,
				...(page ? { page } : {}),
				...(bibliography ? { bibliography } : {}),
				...(item.leader ? { leader: item.leader } : {}),
			});
			paragraphs[0]?.runs.unshift(...item.before);
			paragraphs.at(-1)?.runs.push(...item.after);
			if (instruction && options.heading)
				paragraphs.unshift({
					type: 'paragraph',
					id: '',
					runs: [{ text: options.heading, bold: true }],
				});
			const first = paragraphs[0];
			const last = paragraphs.at(-1);
			// Keep the block identity used by a section boundary and any bookmarks outside the field.
			if (!instruction && first && last) {
				if (paragraphs.length === item.end - item.start + 1)
					for (const [index, paragraph] of paragraphs.entries()) {
						const previous = model.blocks[item.start + index];
						if (previous?.type === 'paragraph') {
							paragraph.id = previous.id;
							if (previous.bookmarks) paragraph.bookmarks = previous.bookmarks;
						}
					}
				const before = model.blocks[item.start];
				const after = model.blocks[item.end];
				if (before?.type === 'paragraph' && first !== last) {
					first.id = before.id;
					if (before.bookmarks) first.bookmarks = before.bookmarks;
				}
				if (after?.type === 'paragraph') {
					last.id = after.id;
					if (after.bookmarks) last.bookmarks = after.bookmarks;
				}
			}
			for (const paragraph of paragraphs) if (!paragraph.id) paragraph.id = newId();
			const section =
				options.columns && instruction
					? references.indexSection(model, item.start, item.end, paragraphs, options.columns, newId)
					: undefined;
			return {
				...item,
				paragraphs: section?.paragraphs ?? paragraphs,
				sections: section?.sections,
			};
		});
	const paginate = kind !== 'BIBLIOGRAPHY';
	const draft = [...model.blocks];
	let draftSections = model.sections;
	if (paginate)
		for (const item of build(() => '1')) {
			draft.splice(item.start, item.end - item.start + 1, ...item.paragraphs);
			draftSections = item.sections ?? draftSections;
		}
	const page = paginate
		? referencePages({
				...model,
				blocks: draft,
				...(draftSections ? { sections: draftSections } : {}),
			})
		: undefined;
	const offsets: number[] = [];
	tr.doc.forEach((_node, offset) => offsets.push(offset));
	offsets.push(tr.doc.content.size);
	for (const item of build(page)) {
		const from = offsets[item.start];
		const to = offsets[item.end + 1];
		if (from === undefined || to === undefined) continue;
		const content = modelToDoc({ ...model, blocks: item.paragraphs }).content;
		if (!tr.doc.slice(from, to).content.eq(content)) tr.replaceWith(from, to, content);
		if (item.sections)
			tr.setDocAttribute('sections', sectionLayoutJson(item.sections)).setDocAttribute(
				'sectionParts',
				sectionPartsJson(item.sections),
			);
	}
	return true;
}

export function selectedCitation(references: WordReferences, view: EditorView) {
	const { $from, $to } = view.state.selection;
	if (!$from.sameParent($to) || $from.parent.type.name !== 'paragraph') return undefined;
	const runs: TextRun[] = [];
	const offsets: number[] = [];
	$from.parent.forEach((node, offset) => {
		const run = inlineNodeRun(node);
		if (run) {
			runs.push(run);
			offsets.push($from.start() + offset);
		}
	});
	offsets.push($from.end());
	for (const field of references.citationSpans(runs)) {
		const from = offsets[field.from];
		const to = offsets[field.to + 1];
		if (
			!field.locked &&
			!field.removed &&
			from !== undefined &&
			to !== undefined &&
			from <= $from.pos &&
			$to.pos <= to
		)
			return { from, to, instruction: field.instr, format: runs[field.runs[0] ?? field.from] };
	}
	return undefined;
}

export function refreshCitationResults(
	references: WordReferences,
	tr: Transaction,
	model: DocumentModel,
	bibliography: Bibliography,
) {
	const evaluate = references.citationEvaluator(bibliography);
	const paragraphs = new Map<
		string,
		Extract<DocumentModel['blocks'][number], { type: 'paragraph' }>
	>();
	for (const block of model.blocks) {
		if (block.type === 'paragraph') paragraphs.set(block.id, block);
		else
			for (const row of block.rows)
				for (const cell of row)
					for (const paragraph of cell.paragraphs) paragraphs.set(paragraph.id, paragraph);
	}
	const changes: { pos: number; node: ProseMirrorNode; next: Fragment }[] = [];
	tr.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		const paragraph = paragraphs.get(String(node.attrs.id));
		if (!paragraph) return false;
		const updated = references.refreshCitations(paragraph, evaluate);
		if (updated !== paragraph)
			changes.push({
				pos,
				node,
				next: Fragment.fromArray(updated.runs.flatMap((run) => runToInlineNodes(run, schema))),
			});
		return false;
	});
	for (const { pos, node, next } of changes.reverse()) {
		const first = node.content.findDiffStart(next);
		const last = node.content.findDiffEnd(next);
		if (first !== null && last) {
			const overlap = Math.max(0, first - Math.min(last.a, last.b));
			tr.replaceWith(
				pos + 1 + first,
				pos + 1 + last.a + overlap,
				next.cut(first, last.b + overlap),
			);
		}
	}
}

export function selectedRunFormat(
	view: EditorView,
): Pick<TextRun, 'bold' | 'italic' | 'underline'> {
	const { from, to, $from } = view.state.selection;
	let run: TextRun | undefined;
	view.state.doc.nodesBetween(from, Math.max(to, Math.min(from + 1, $from.end())), (node) => {
		if (node.isText && !run) run = inlineNodeRun(node);
		return run === undefined;
	});
	return {
		...(run?.bold ? { bold: true } : {}),
		...(run?.italic ? { italic: true } : {}),
		...(run?.underline ? { underline: true } : {}),
	};
}
