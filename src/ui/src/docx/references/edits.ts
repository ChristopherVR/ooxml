import { closeHistory } from 'prosemirror-history';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	sectionLayoutJson,
	styleCatalogsOf,
	type DocumentModel,
	type TextRun,
} from 'ooxml-core/docx';
import {
	inlineNodeRun,
	newControlId,
	runToInlineNodes,
	sectionPartsJson,
	type Bibliography,
	type IndexLeader,
	type ReferencePage,
	type ReferenceTable,
	type WordReferences,
} from 'ooxml-core/docx/ui';
import { withBibliographySources } from '../document-information-attr';
import { modelToDoc } from '../model-adapter';
import { schema } from '../schema';
import { markAllPositions } from './mark-all';
import type { EditorLocale } from '../localization';
import { referencesText } from './strings';
import { referencePages } from './reference-pages';

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
	else for (const at of markAllPositions(tr.doc, all).reverse()) inlineReference(tr, at, at, runs);
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
	const xml = references.writeBibliographyXml(data, original?.xml);
	if (tr.doc.attrs.bibliographySources !== xml) tr.setDocAttribute('bibliographySources', xml);
	return withBibliographySources(model, xml, original?.part);
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
		locale: EditorLocale;
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
	const galleryHeading = instruction && kind === 'BIBLIOGRAPHY' ? options.heading : undefined;
	const galleryStyles = galleryHeading ? references.bibliographyGalleryStyles(model) : undefined;
	if (galleryStyles && galleryStyles !== model.paragraphStyles) {
		model = { ...model, paragraphStyles: galleryStyles };
		tr.setDocAttribute('styleCatalogs', styleCatalogsOf(model));
	}
	const galleryId = galleryHeading ? newControlId(tr.doc) : undefined;
	const newId = () => `ref-${references.createOfficeGuid()}`;
	const ordered = [...edit].sort((a, b) => b.start - a.start);
	const build = (page?: ReferencePage) =>
		ordered.map((item) => {
			const paragraphs = references.buildReferenceTable(model, item.instruction, {
				empty,
				locale: options.locale,
				translateCategory: (name) => referencesText(options.locale, name),
				blockIndex: item.start,
				...(page ? { page } : {}),
				...(bibliography ? { bibliography } : {}),
				...(item.leader ? { leader: item.leader } : {}),
			});
			paragraphs[0]?.runs.unshift(...item.before);
			paragraphs.at(-1)?.runs.push(...item.after);
			if (galleryId && galleryHeading) {
				references.bibliographyGallery(model, galleryId, galleryHeading, paragraphs);
				if (item.end >= model.blocks.length - 1)
					paragraphs.push({ type: 'paragraph', id: '', runs: [] });
			}
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
					if (before.contentControl) {
						// The last paragraph may belong to a different control or be outside it.
						for (const paragraph of paragraphs.slice(0, -1))
							paragraph.contentControl = before.contentControl;
						if (after?.type === 'paragraph' && after.contentControl)
							last.contentControl = after.contentControl;
					}
				}
				if (after?.type === 'paragraph') {
					// A separate end paragraph can carry a section break and following body text.
					// Keep its paragraph properties when refreshing the field result.
					if (after.runs[0]?.fieldChar === 'end') Object.assign(last, after, { runs: last.runs });
					else {
						last.id = after.id;
						if (after.bookmarks) last.bookmarks = after.bookmarks;
					}
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
