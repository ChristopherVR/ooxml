import { styleCatalogsOf } from 'ooxml-core/docx';
import { withStyleCatalogs } from './style-catalog-state';
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
	PARAGRAPH_EAST_ASIAN_KEYS,
	PARAGRAPH_KEEP_KEYS as KEEP_KEYS,
} from 'ooxml-core/docx/ui';
import { sameJson, sameRuns } from 'ooxml-core/docx/ui';
import { appendInlineNode, runToInlineNodes, type NoteNumberLookup } from './run-adapter';
import {
	convertStructuredTable,
	convertSimpleTable,
	tableControl,
	tableNode,
	tableFormattingFromNode,
	tablePropertiesFromNode,
} from './table-model-adapter';
import { sectionLayoutJson, sectionsFromLayout } from 'ooxml-core/docx';
import { sectionsOf } from './section-commands';
import { sectionPartsJson, restoreSectionParts } from 'ooxml-core/docx/ui';
import { notePartsJson, restoreNoteParts } from 'ooxml-core/docx/ui';
import { commentThreadsFromDoc } from 'ooxml-core/docx/ui';
import {
	documentInformationJson,
	readDocumentInformation,
	withBibliographySources,
} from './document-information-attr';
import { signedTwips } from 'ooxml-core/units';

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
			model.theme,
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
			noteParts: notePartsJson(model),
			commentThreads: JSON.stringify(model.comments ?? []),
			documentInformation: documentInformationJson(model),
			mailMerge: model.mailMerge ?? null,
			bibliographySources: model.packageInformation?.bibliography?.xml ?? null,
			styleCatalogs: styleCatalogsOf(model),
			documentTheme: model.theme ?? null,
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

/**
 * `model` with the properties and package information of the document attribute, when the
 * transaction changed it (Inspect Document's Remove All, or its undo). Otherwise the model's own
 * fields stand, so a File > Info property edit made on the model is kept.
 */
function withDocumentInformation<T extends Partial<DocumentModel>>(
	model: T,
	value: unknown,
	previous: unknown,
): T {
	if (value === previous) return model;
	const info = readDocumentInformation(value);
	if (!info) return model;
	const { properties: _properties, packageInformation: _info, ...rest } = model;
	return { ...rest, ...info } as T;
}

export function docToModel(
	doc: ReturnType<typeof modelToDoc>,
	prior: DocumentModel,
	previousDoc?: ProseMirrorNode,
): DocumentModel {
	let nextId = 0;
	const previousParagraphs = new Map<string, Paragraph>();
	const priorTables = new Map<string, Table>();
	const duplicateIds = new Set<string>();
	const remember = (paragraph: Paragraph) => {
		if (previousParagraphs.has(paragraph.id)) duplicateIds.add(paragraph.id);
		previousParagraphs.set(paragraph.id, paragraph);
	};
	const rememberTable = (table: Table): void => {
		for (const row of table.rows)
			for (const cell of row) {
				cell.paragraphs.forEach(remember);
				for (const block of cell.sourceBlocks ?? [])
					if (block.type === 'table') rememberTable(block);
			}
	};
	for (const block of prior.blocks) {
		if (block.type === 'paragraph') remember(block);
		else {
			if (priorTables.has(block.id)) duplicateIds.add(block.id);
			priorTables.set(block.id, block);
			rememberTable(block);
		}
	}

	const unchanged = new WeakMap<ProseMirrorNode, Paragraph>();
	const unchangedTables = new WeakMap<ProseMirrorNode, Table>();
	previousDoc?.descendants((node) => {
		if (duplicateIds.has(String(node.attrs.id))) return true;
		if (node.type.name === 'paragraph') {
			const paragraph = previousParagraphs.get(String(node.attrs.id));
			if (paragraph) unchanged.set(node, paragraph);
			return false;
		}
		if (node.type.name === 'table') {
			const table = priorTables.get(String(node.attrs.id));
			if (table) unchangedTables.set(node, table);
		}
		return true;
	});
	const convertParagraph = (node: typeof doc): Paragraph => {
		const cached = unchanged.get(node);
		if (cached) return cached;
		const runs: TextRun[] = [];
		node.forEach((child) => appendInlineNode(runs, child));
		if (!runs.length) runs.push({ text: '' });
		const id = String(node.attrs.id || `p-edit-${++nextId}`);
		const previous = previousParagraphs.get(id);
		if (
			previous &&
			(previous.sourceParagraphPropertiesXml ?? null) ===
				(node.attrs.sourceParagraphPropertiesXml ?? null) &&
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
			(previous.pageBreakBefore ?? null) === node.attrs.pageBreakBefore &&
			sameJson(previous.tabStops ?? null, node.attrs.tabStops ?? null) &&
			[...KEEP_KEYS, ...PARAGRAPH_EAST_ASIAN_KEYS].every(
				(key) => (previous[key] ?? null) === (node.attrs[key] ?? null),
			) &&
			sameJson(previous.dropCap ?? null, node.attrs.dropCap ?? null) &&
			sameJson(previous.frame ?? null, node.attrs.frame ?? null) &&
			sameJson(previous.borders ?? null, node.attrs.borders ?? null) &&
			JSON.stringify(previous.percentageShading ?? null) ===
				JSON.stringify(node.attrs.percentageShading ?? null) &&
			(previous.shadingFill ?? null) === (node.attrs.shadingFill ?? null) &&
			sameJson(previous.bookmarks ?? [], node.attrs.bookmarks ?? []) &&
			sameJson(previous.permissions ?? null, node.attrs.permissions ?? null) &&
			sameJson(previous.contentControl ?? null, node.attrs.contentControl ?? null)
		)
			return previous;
		return paragraphFromAttrs(node.attrs, id, runs);
	};

	const blocks: Block[] = [];
	doc.forEach((node) => {
		if (node.type.name === 'paragraph') blocks.push(convertParagraph(node as typeof doc));
		if (node.type.name === 'table') {
			const cached = unchangedTables.get(node);
			if (cached) {
				blocks.push(cached);
				return;
			}
			const id = String(node.attrs.id || `t-edit-${++nextId}`);
			const prior = priorTables.get(id);
			const asParagraph = (n: unknown) => convertParagraph(n as typeof doc);
			if (
				prior &&
				(prior.structureEditable === false ||
					(node.attrs.structureEditable === false && node.attrs.restructurable))
			)
				blocks.push(convertStructuredTable(node, prior, asParagraph));
			else {
				const {
					style: priorStyle,
					look: _look,
					layout: _layout,
					grid: _grid,
					widthTwips: _width,
					widthPercent: _percent,
					borders: _borders,
					structureEditable: _structureEditable,
					restructurable: _restructurable,
					...descriptors
				} = prior ?? {};
				const style =
					typeof node.attrs.tableStyleId === 'string' ? node.attrs.tableStyleId : priorStyle;
				blocks.push({
					...(prior
						? {
								...descriptors,
								contentControl: undefined,
							}
						: {
								// Indent and alignment have no table-tab control; read them from the node.
								...(typeof node.attrs.indentTwips === 'number' &&
									Number.isSafeInteger(node.attrs.indentTwips) && {
										indentTwips: signedTwips(node.attrs.indentTwips),
									}),
								...(['left', 'center', 'right'].includes(node.attrs.alignment) && {
									alignment: node.attrs.alignment,
								}),
							}),
					type: 'table',
					id,
					...(style && { style }),
					...tablePropertiesFromNode(node, prior),
					rows: convertSimpleTable(node, asParagraph, prior),
					...tableFormattingFromNode(node),
					...(prior?.structureEditable !== undefined || node.attrs.structureEditable === false
						? { structureEditable: node.attrs.structureEditable === true }
						: {}),
					...(node.attrs.restructurable ? { restructurable: true } : {}),
					...tableControl(node, prior),
				} as Table);
			}
		}
	});
	// The editor document is the source of truth for section layout (so undo covers page setup).
	const {
		sections: priorSections,
		footnotes: priorFootnotes,
		endnotes: priorEndnotes,
		comments: priorComments,
		evenAndOddHeaders: _evenOdd,
		trackChanges: _trackChanges,
		trackFormatting: _trackFormatting,
		trackMoves: _trackMoves,
		pageColor: _pageColor,
		autoHyphenation: _hyphenation,
		mailMerge: priorMailMerge,
		...kept
	} = prior;
	const documentInformation = previousDoc
		? withDocumentInformation(
				kept,
				doc.attrs.documentInformation,
				previousDoc.attrs.documentInformation,
			)
		: kept;
	const rest = withBibliographySources(
		documentInformation,
		doc.attrs.bibliographySources,
		prior.packageInformation?.bibliography?.part,
	);
	const sectionLayout =
		typeof doc.attrs.sections === 'string'
			? sectionsFromLayout(doc.attrs.sections, priorSections, blocks)
			: keepHeaderFooterSections(priorSections, blocks, doc).sections;
	const sections =
		typeof doc.attrs.sectionParts === 'string' && sectionLayout
			? restoreSectionParts(sectionLayout, doc.attrs.sectionParts, blocks)
			: sectionLayout;
	const threads = commentThreadsFromDoc(doc);
	const mailMerge: DocumentModel['mailMerge'] = previousDoc
		? (doc.attrs.mailMerge ?? undefined)
		: priorMailMerge;
	return {
		...(previousDoc && previousDoc.attrs.styleCatalogs !== doc.attrs.styleCatalogs
			? withStyleCatalogs(rest, doc)
			: rest),
		...(threads === undefined
			? priorComments !== undefined
				? { comments: priorComments }
				: {}
			: threads.length
				? { comments: threads }
				: {}),
		...restoreNoteParts(doc.attrs.noteParts, {
			...(priorFootnotes ? { footnotes: priorFootnotes } : {}),
			...(priorEndnotes ? { endnotes: priorEndnotes } : {}),
		}),
		blocks,
		...(mailMerge && { mailMerge }),
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
