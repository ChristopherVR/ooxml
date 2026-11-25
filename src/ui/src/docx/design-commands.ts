import { editDocumentDesign, styleCatalogsOf, type DesignEdit } from 'ooxml-core/docx';
import { closeHistory } from 'prosemirror-history';
import type { EditorCore } from './editor-core';
import { canEditDesign } from './design-ribbon-controls';

export function applyDesignEdit(core: EditorCore, edit: DesignEdit): boolean {
	const view = core.view;
	if (!view || !canEditDesign(core)) return false;
	const model = core.model;
	const next = editDocumentDesign(model, edit);
	const key = edit.kind === 'spacing' ? 'styleCatalogs' : 'documentTheme';
	const value = edit.kind === 'spacing' ? styleCatalogsOf(next) : next.theme;
	if (JSON.stringify(value) === JSON.stringify(view.state.doc.attrs[key])) return false;
	view.dispatch(closeHistory(view.state.tr).setDocAttribute(key, value));
	return true;
}
