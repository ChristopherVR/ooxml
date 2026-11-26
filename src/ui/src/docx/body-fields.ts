import type { Node } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { fieldName } from 'ooxml-core/docx';
import { loadWordReferences } from 'ooxml-core/docx/ui';
import { refreshFieldResults } from './field-update';
import type { InsertControllerHost } from './insert-controller';
import { translate, type EditorLocale } from './localization';
import { docToModel } from './model-adapter';
import { blockPageNumbers, refreshTableOfContents } from './toc-commands';

const REFERENCE_FIELDS = new Set(['INDEX', 'TOA', 'BIBLIOGRAPHY', 'CITATION']);

function hasReferences(doc: Node): boolean {
	let found = false;
	const codes: string[] = [];
	doc.descendants((node) => {
		if (found) return false;
		if (node.type.name === 'fieldMarker') {
			const { kind, code } = node.attrs;
			if (kind === 'begin') codes.push('');
			else if (kind === 'code') {
				const last = codes.length - 1;
				const prior = codes[last];
				if (prior !== undefined && prior.length <= 8192) codes[last] = prior + String(code ?? '');
			} else if (kind === 'separate' || kind === 'end') {
				const instruction = kind === 'end' ? codes.pop() : codes.at(-1);
				found = REFERENCE_FIELDS.has(fieldName(instruction ?? ''));
			}
		} else if (node.isInline) {
			found = node.marks.some(
				(mark) =>
					mark.type.name === 'field' && REFERENCE_FIELDS.has(fieldName(String(mark.attrs.instr))),
			);
		}
		return !found;
	});
	return found;
}

export async function updateBodyFields(
	view: EditorView,
	host: Pick<InsertControllerHost, 'view' | 'model' | 'reportError'>,
	locale: EditorLocale,
): Promise<void> {
	try {
		const original = view.state.doc;
		const loaded = hasReferences(original)
			? await Promise.all([import('./references/controller'), loadWordReferences()])
			: undefined;
		if (host.view() !== view || original !== view.state.doc || !view.editable) {
			host.reportError(new Error(translate(locale, 'fieldUpdate.changed')));
			return;
		}
		let model = host.model();
		let numbers: ReturnType<typeof blockPageNumbers> | undefined;
		const tr = view.state.tr;
		let checkpoint = tr.doc;
		const readEdits = () => {
			if (checkpoint === tr.doc) return;
			model = docToModel(tr.doc, model, checkpoint);
			checkpoint = tr.doc;
		};
		refreshFieldResults(tr, (id) => (numbers ??= blockPageNumbers(model)).get(id), false, model);
		if (loaded) {
			readEdits();
			const [{ refreshReferences }, references] = loaded;
			refreshReferences(references, tr, model, locale);
		}
		for (const label of [undefined, 'Figure', 'Table', 'Equation']) {
			readEdits();
			refreshTableOfContents(tr, model, undefined, label);
		}
		if (tr.docChanged) view.dispatch(closeHistory(tr));
		else host.reportError(new Error(translate(locale, 'fieldUpdate.none')));
	} catch (error) {
		host.reportError(error instanceof Error ? error : new Error(String(error)));
	}
}
