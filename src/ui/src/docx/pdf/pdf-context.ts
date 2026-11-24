/** What the editor can offer the PDF Options step; kept apart from the lazily loaded PDF writer. */
import type { EditorCore } from '../editor-core';
import type { PdfOptionsContext } from './pdf-options-form';

/** Print Layout's current page and page count (as its status bar shows them); undefined in Draft. */
export function editorPageStatus(core: EditorCore): { current: number; total: number } | undefined {
	if (core.pages.viewMode !== 'print') return undefined;
	return core.shell.printLayout?.pageStatus() ?? undefined;
}

export function editorPdfContext(core: EditorCore): PdfOptionsContext {
	const selection = core.view?.state.selection;
	return {
		pages: editorPageStatus(core),
		hasSelection: Boolean(selection && !selection.empty),
	};
}
