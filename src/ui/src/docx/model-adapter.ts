import { EditorState, Transaction } from 'prosemirror-state';
import type {
	DocumentModel,
	Block,
	Paragraph,
	SectionProperties,
	Table,
	TextRun,
} from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	computeListLabels,
	displayListLabel,
	formatNoteNumber,
	numberNotesInOrder,
} from 'ooxml-core/docx';
import { schema } from './schema';
import {
	paragraphAttrs,
	paragraphFromAttrs,
	PARAGRAPH_KEEP_KEYS as KEEP_KEYS,
} from 'ooxml-core/docx/ui';
import { sameJson, sameRuns } from 'ooxml-core/docx/ui';
import { appendInlineNode, runToInlineNodes, type NoteNumberLookup } from './run-adapter';
import {
	convertMergedTable,
	convertSimpleTable,
	tableNode,
	tableFormattingFromNode,
} from './table-model-adapter';
import { parseBordersJson } from 'ooxml-core/docx/ui';
import { sectionLayoutJson, sectionsFromLayout } from 'ooxml-core/docx';
import { sectionsOf } from './section-commands';
import { sectionPartsJson, restoreSectionParts } from 'ooxml-core/docx/ui';

type ListLabels = ReturnType<typeof computeListLabels>;

function paragraphNode(paragraph: Paragraph, labels: ListLabels, noteNumber?: NoteNumberLookup) {
	const label = labels.get(paragraph.id);
	const children = paragraph.runs.flatMap((run) => runToInlineNodes(run, noteNumber));
	return schema.node(
		'paragraph',
		{
			...paragraphAttrs(paragraph),
			listLabelText: label ? displayListLabel(label.text) : null,
			listSuffix: label?.suffix ?? null,
			listIndentLeftTwips: label?.indentLeftTwips ?? null,
			listHangingTwips: label?.hangingTwips ?? null,
			listFirstLineTwips: label?.firstLineTwips ?? null,
		},
		children,
	);
}

export function modelToDoc(model: DocumentModel) {
	const labels = computeListLabels(model);
	const footnoteOrder = numberNotesInOrder(model.blocks, 'footnote');
	const endnoteOrder = numberNotesInOrder(model.blocks, 'endnote');
	const noteNumber: NoteNumberLookup = (kind, id) => {
		const number = (kind === 'footnote' ? footnoteOrder : endnoteOrder).get(id) ?? 1;
		const format =
			kind === 'footnote'
				? (model.footnoteNumFmt ?? 'decimal')
				: (model.endnoteNumFmt ?? 'lowerRoman');
		return { number, label: formatNoteNumber(number, format) };
	};
	const blocks = model.blocks.map((block) => {
		if (block.type === 'paragraph') return paragraphNode(block, labels, noteNumber);
		return tableNode(
			block,
			(paragraph) => paragraphNode(paragraph, labels, noteNumber),
			model.tableStyles,
		);
	});
	return schema.node(
		'doc',
		{
			pageWidth: model.page.width,
			pageHeight: model.page.height,
			marginTop: model.page.marginTop,
			marginRight: model.page.marginRight,
			marginBottom: model.page.marginBottom,
			marginLeft: model.page.marginLeft,
			sections: model.sections ? sectionLayoutJson(model.sections) : null,
			sectionParts: sectionPartsJson(model.sections),
			evenAndOddHeaders: Boolean(model.evenAndOddHeaders),
			trackChanges: Boolean(model.trackChanges),
			trackFormatting: model.trackFormatting !== false,
			trackMoves: model.trackMoves !== false,
			pageColor: model.pageColor ?? null,
			autoHyphenation: Boolean(model.autoHyphenation),
			footnoteNumFmt: model.footnoteNumFmt ?? null,
			endnoteNumFmt: model.endnoteNumFmt ?? null,
		},
		blocks,
	);
}

/** A table inserted in the editor keeps the borders it was created with (see insertTable). */
function tableBordersFromNode(value: unknown): Partial<Table> {
	const borders = parseBordersJson(value);
	return borders ? { borders: borders as NonNullable<Table['borders']> } : {};
}
/**
 * A document with no recorded section layout (a new document nobody changed page setup on) has no
 * sections of its own, but header/footer content added since must survive: one section is
 * synthesized from the page geometry and keeps that content.
 */
function keepHeaderFooterSections(
	prior: SectionProperties[] | undefined,
	blocks: Block[],
	doc: ProseMirrorNode,
): { sections?: SectionProperties[] } {
	const source = prior?.find((section) => section.headers || section.footers);
	if (!source) return {};
	const page = {
		width: doc.attrs.pageWidth,
		height: doc.attrs.pageHeight,
		marginTop: doc.attrs.marginTop,
		marginRight: doc.attrs.marginRight,
		marginBottom: doc.attrs.marginBottom,
		marginLeft: doc.attrs.marginLeft,
	};
	const [section] = sectionsOf({ blocks, page } as DocumentModel);
	if (!section) return {};
	return {
		sections: [
			{
				...section,
				...(source.headers ? { headers: source.headers } : {}),
				...(source.footers ? { footers: source.footers } : {}),
			},
		],
	};
}

export function docToModel(
	doc: ReturnType<typeof modelToDoc>,
	prior: DocumentModel,
): DocumentModel {
	let nextId = 0;
	const previousParagraphs = new Map<string, Paragraph>();
	const priorTables = new Map<string, Table>();
	const remember = (paragraph: Paragraph) => previousParagraphs.set(paragraph.id, paragraph);
	for (const block of prior.blocks) {
		if (block.type === 'paragraph') remember(block);
		else {
			priorTables.set(block.id, block);
			for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(remember);
		}
	}

	const convertParagraph = (node: typeof doc): Paragraph => {
		const runs: TextRun[] = [];
		node.forEach((child) => appendInlineNode(runs, child));
		if (!runs.length) runs.push({ text: '' });
		const id = String(node.attrs.id || `p-edit-${++nextId}`);
		const previous = previousParagraphs.get(id);
		if (
			previous &&
			(previous.restoredParagraphPropertiesXml ?? null) ===
				(node.attrs.restoredParagraphPropertiesXml ?? null) &&
			sameJson(previous.markRevision ?? null, node.attrs.markRevision ?? null) &&
			sameJson(previous.formatRevision ?? null, node.attrs.formatRevision ?? null) &&
			sameRuns(previous.runs, runs) &&
			previous.align === (node.attrs.align ?? undefined) &&
			previous.justification === (node.attrs.justification ?? undefined) &&
			previous.outlineLevel === (node.attrs.outlineLevel ?? undefined) &&
			previous.direction === (node.attrs.direction ?? undefined) &&
			(previous.style || '') === (node.attrs.style || '') &&
			previous.spacingBeforeTwips === (node.attrs.spacingBeforeTwips ?? undefined) &&
			previous.spacingAfterTwips === (node.attrs.spacingAfterTwips ?? undefined) &&
			previous.lineSpacingTwips === (node.attrs.lineSpacingTwips ?? undefined) &&
			previous.lineSpacingRule === (node.attrs.lineSpacingRule ?? undefined) &&
			previous.indentLeftTwips === (node.attrs.indentLeftTwips ?? undefined) &&
			previous.indentRightTwips === (node.attrs.indentRightTwips ?? undefined) &&
			previous.indentStartTwips === (node.attrs.indentStartTwips ?? undefined) &&
			previous.indentEndTwips === (node.attrs.indentEndTwips ?? undefined) &&
			previous.firstLineTwips === (node.attrs.firstLineTwips ?? undefined) &&
			previous.hangingTwips === (node.attrs.hangingTwips ?? undefined) &&
			(previous.numbering?.numId ?? null) === (node.attrs.numId ?? null) &&
			(previous.numbering ? previous.numbering.level : null) === (node.attrs.ilvl ?? null) &&
			Boolean(previous.pageBreakBefore) === Boolean(node.attrs.pageBreakBefore) &&
			sameJson(previous.tabStops ?? null, node.attrs.tabStops ?? null) &&
			KEEP_KEYS.every((key) => (previous[key] ?? null) === (node.attrs[key] ?? null)) &&
			sameJson(previous.dropCap ?? null, node.attrs.dropCap ?? null) &&
			sameJson(previous.borders ?? null, node.attrs.borders ?? null) &&
			(previous.shadingFill ?? null) === (node.attrs.shadingFill ?? null) &&
			sameJson(previous.bookmarks ?? [], node.attrs.bookmarks ?? [])
		)
			return previous;
		return paragraphFromAttrs(node.attrs, id, runs);
	};

	const blocks: Block[] = [];
	doc.forEach((node) => {
		if (node.type.name === 'paragraph') blocks.push(convertParagraph(node as typeof doc));
		if (node.type.name === 'table') {
			const id = String(node.attrs.id || `t-edit-${++nextId}`);
			const prior = priorTables.get(id);
			const asParagraph = (n: unknown) => convertParagraph(n as typeof doc);
			if (prior && prior.structureEditable === false)
				blocks.push(convertMergedTable(node, prior, asParagraph));
			else
				blocks.push({
					// Table-level properties (grid, width, borders, style, look) have no editor node;
					// carry them over from the prior model so editing cell text never drops them.
					...(prior
						? {
								...prior,
								structureEditable: undefined,
							}
						: tableBordersFromNode(node.attrs.borders)),
					type: 'table',
					id,
					rows: convertSimpleTable(node, asParagraph, prior),
					...tableFormattingFromNode(node),
					...(node.attrs.structureEditable === false ? { structureEditable: false } : {}),
				} as Table);
		}
	});
	// The editor document is the source of truth for section layout (so undo covers page setup).
	const {
		sections: priorSections,
		evenAndOddHeaders: _evenOdd,
		trackChanges: _trackChanges,
		trackFormatting: _trackFormatting,
		trackMoves: _trackMoves,
		pageColor: _pageColor,
		autoHyphenation: _hyphenation,
		...rest
	} = prior;
	const sectionLayout =
		typeof doc.attrs.sections === 'string'
			? sectionsFromLayout(doc.attrs.sections, priorSections, blocks)
			: keepHeaderFooterSections(priorSections, blocks, doc).sections;
	const sections =
		typeof doc.attrs.sectionParts === 'string' && sectionLayout
			? restoreSectionParts(sectionLayout, doc.attrs.sectionParts, blocks)
			: sectionLayout;
	return {
		...rest,
		blocks,
		...(doc.attrs.evenAndOddHeaders ? { evenAndOddHeaders: true } : {}),
		trackChanges: Boolean(doc.attrs.trackChanges),
		trackFormatting: doc.attrs.trackFormatting !== false,
		trackMoves: doc.attrs.trackMoves !== false,
		...(doc.attrs.pageColor ? { pageColor: String(doc.attrs.pageColor) } : {}),
		...(doc.attrs.autoHyphenation ? { autoHyphenation: true } : {}),
		...(sections ? { sections } : {}),
		page: {
			width: doc.attrs.pageWidth,
			height: doc.attrs.pageHeight,
			marginTop: doc.attrs.marginTop,
			marginRight: doc.attrs.marginRight,
			marginBottom: doc.attrs.marginBottom,
			marginLeft: doc.attrs.marginLeft,
		},
	};
}

export function assignMissingParagraphIds(state: EditorState): Transaction | null {
	const reserved = new Set<string>();
	const seen = new Set<string>();
	state.doc.descendants((node) => {
		if (node.type.name === 'paragraph' && node.attrs.id) reserved.add(String(node.attrs.id));
	});
	let next = 1;
	let transaction = state.tr;
	state.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return;
		const id = String(node.attrs.id || '');
		if (id && !seen.has(id)) {
			seen.add(id);
			return;
		}
		let generated = '';
		do {
			generated = `p-edit-${next++}`;
		} while (reserved.has(generated));
		reserved.add(generated);
		seen.add(generated);
		transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, id: generated });
	});
	return transaction.docChanged ? transaction : null;
}
