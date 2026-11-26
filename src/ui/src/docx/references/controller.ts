import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from 'ooxml-core/docx';
import type { Bibliography, ReferenceTable, WordReferences } from 'ooxml-core/docx/ui';
import type { EditorCore } from '../editor-core';
import { referencesText, type ReferencesText } from './strings';
import { docToModel } from '../model-adapter';
import { REFERENCE_LABELS, type ReferencesCommand } from '../references-ribbon';
import { referencesDialog } from './dialog';
import {
	bibliographyOf,
	changeReferenceTables,
	dispatchReference,
	inlineReference,
	markEntry,
	refreshCitationResults,
	selectedCitation,
	selectedRunFormat,
	setBibliography,
} from './edits';
import {
	authorityDialog,
	citationOptions,
	insertIndexDialog,
	markAuthorityDialog,
	markIndexDialog,
} from './entry-dialogs';
import { insertCitationDialog, manageSourcesDialog, styleDialog } from './source-dialogs';

const EMPTY: Record<ReferenceTable, ReferencesText> = {
	INDEX: 'No index entries found.',
	TOA: 'No table of authorities entries found.',
	BIBLIOGRAPHY: 'No bibliography sources found.',
};
const MISSING: Record<ReferenceTable, ReferencesText> = {
	INDEX: 'This document has no index to update.',
	TOA: 'This document has no table of authorities to update.',
	BIBLIOGRAPHY: 'This document has no bibliography to update.',
};

export function refreshReferences(
	references: WordReferences,
	tr: Transaction,
	model: DocumentModel,
	locale: EditorCore['locale'],
): void {
	let current = model;
	for (const kind of ['INDEX', 'TOA', 'BIBLIOGRAPHY'] as const) {
		const bibliography = kind === 'BIBLIOGRAPHY' ? bibliographyOf(references, current) : undefined;
		if (kind === 'BIBLIOGRAPHY' && !bibliography) continue;
		const checkpoint = tr.doc;
		changeReferenceTables(references, tr, current, {
			kind,
			empty: referencesText(locale, EMPTY[kind]),
			...(bibliography ? { bibliography } : {}),
		});
		if (checkpoint !== tr.doc) current = docToModel(tr.doc, current, checkpoint);
		if (bibliography) refreshCitationResults(references, tr, current, bibliography);
	}
}

export function runReferences(
	references: WordReferences,
	core: EditorCore,
	view: EditorView,
	command: ReferencesCommand,
): void {
	const {
		REFERENCE_LCID,
		authorityCode,
		authorityEntryCode,
		citationCode,
		citationEvaluator,
		editCitationCode,
		indexCode,
		indexEntryCode,
		readCitation,
		referenceFieldRuns,
	} = references;
	const { model, locale } = core;
	const t = (key: ReferencesText) => referencesText(locale, key);
	const selection = view.state.selection;
	const text = view.state.doc.textBetween(selection.from, selection.to, '\n').trim();
	const singleParagraph =
		selection.$from.sameParent(selection.$to) && selection.$from.parent.type.name === 'paragraph';
	const open = () => referencesDialog(core, REFERENCE_LABELS[command]);
	const table = (
		kind: ReferenceTable,
		options: Partial<Parameters<typeof changeReferenceTables>[3]> = {},
	) => {
		const tr = view.state.tr;
		if (!changeReferenceTables(references, tr, model, { ...options, kind, empty: t(EMPTY[kind]) }))
			core.host.reportError(new Error(t(MISSING[kind])));
		dispatchReference(view, tr);
	};
	switch (command) {
		case 'markEntry': {
			if (!singleParagraph) break;
			const bookmarks = new Set<string>();
			for (const block of model.blocks) {
				if (block.type === 'paragraph')
					for (const name of block.bookmarks ?? []) bookmarks.add(name);
				else
					for (const row of block.rows)
						for (const cell of row)
							for (const p of cell.paragraphs)
								for (const name of p.bookmarks ?? []) bookmarks.add(name);
			}
			markIndexDialog(open(), text.slice(0, 512), [...bookmarks].slice(0, 1000), (options, all) =>
				markEntry(references, view, indexEntryCode(options), all ? text.slice(0, 512) : undefined),
			);
			return;
		}
		case 'markAuthority':
			if (!singleParagraph) break;
			markAuthorityDialog(open(), text.slice(0, 512), (long, short, category) =>
				markEntry(
					references,
					view,
					authorityEntryCode(long, short, category),
					undefined,
					selectedRunFormat(view),
				),
			);
			return;
		case 'insertIndex':
			insertIndexDialog(references, open(), locale, (options) =>
				table('INDEX', {
					instruction: indexCode(options),
					leader: options.leader,
					columns: options.columns,
				}),
			);
			return;
		case 'insertAuthority':
			authorityDialog(open(), (category, passim, keep) =>
				table('TOA', { instruction: authorityCode(category, passim, keep) }),
			);
			return;
		case 'updateIndex':
			table('INDEX');
			return;
		case 'updateAuthority':
			table('TOA');
			return;
		default: {
			const bibliography = bibliographyOf(references, model);
			if (!bibliography) {
				core.host.reportError(new Error(t('references.sourceListUnreadable')));
				return;
			}
			const syncSources = (data: Bibliography) => {
				const tr = view.state.tr;
				const next = setBibliography(references, tr, model, data);
				changeReferenceTables(references, tr, next, {
					kind: 'BIBLIOGRAPHY',
					empty: t(EMPTY.BIBLIOGRAPHY),
					bibliography: data,
				});
				refreshCitationResults(references, tr, next, data);
				dispatchReference(view, tr);
			};
			if (command === 'sources')
				manageSourcesDialog(references, open(), bibliography, (sources) =>
					syncSources({ ...bibliography, sources }),
				);
			else if (command === 'style')
				styleDialog(open(), bibliography.style, (style) => syncSources({ ...bibliography, style }));
			else if (command === 'bibliography' || command === 'worksCited')
				table('BIBLIOGRAPHY', {
					instruction: 'BIBLIOGRAPHY',
					heading: t(REFERENCE_LABELS[command]),
					bibliography,
				});
			else if (command === 'updateBibliography') {
				const tr = view.state.tr;
				const found = changeReferenceTables(references, tr, model, {
					kind: 'BIBLIOGRAPHY',
					empty: t(EMPTY.BIBLIOGRAPHY),
					bibliography,
				});
				refreshCitationResults(references, tr, model, bibliography);
				if (!found && !tr.docChanged) core.host.reportError(new Error(t(MISSING.BIBLIOGRAPHY)));
				dispatchReference(view, tr);
			} else if (command === 'citation' && singleParagraph) {
				insertCitationDialog(references, open(), bibliography, (tag, options, placeholder) => {
					const tr = view.state.tr;
					const data = placeholder
						? { ...bibliography, sources: [...bibliography.sources, placeholder] }
						: bibliography;
					if (placeholder) setBibliography(references, tr, model, data);
					const instruction = citationCode(tag, REFERENCE_LCID[locale] ?? 1033, options);
					const result = citationEvaluator(data)(instruction) ?? '';
					inlineReference(
						tr,
						selection.to,
						selection.to,
						referenceFieldRuns(instruction, [{ text: result }]),
					);
					dispatchReference(view, tr);
				});
			} else if (command === 'editCitation') {
				const citation = selectedCitation(references, view);
				const options = citation && readCitation(citation.instruction);
				if (citation && options) {
					const dialog = open();
					const read = citationOptions(dialog, options);
					dialog.show(
						dialog.action(
							'OK',
							() => {
								const instruction = editCitationCode(citation.instruction, read());
								const result = instruction && citationEvaluator(bibliography)(instruction);
								if (!instruction || result === undefined)
									throw new Error(t('references.citationSourceMissing'));
								const tr = view.state.tr;
								inlineReference(
									tr,
									citation.from,
									citation.to,
									referenceFieldRuns(instruction, [{ ...citation.format, text: result }]),
								);
								dispatchReference(view, tr);
							},
							true,
						),
					);
				} else core.host.reportError(new Error(t('references.selectCitation')));
			} else break;
			return;
		}
	}
	core.host.reportError(new Error(t('references.paragraph')));
}
