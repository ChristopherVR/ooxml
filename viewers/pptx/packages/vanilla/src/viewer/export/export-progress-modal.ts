import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxUiExportProgressElement } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';

/**
 * ExportProgressModal: the progress modal shown while a multi-slide export (PDF / GIF / WebM)
 * runs. A thin adapter over the shared `pptx-ui-export-progress`, which owns the dialog and
 * progress bar semantics, focus and the look, and is deliberately not dismissable: the export
 * ends by completing, erroring or Cancel. `onCancel` aborts cooperatively (the export loops
 * check the signal between slides).
 */
export interface ExportProgressModal {
	/** Show the modal with a fresh title, zeroed bar, and initial status. */
	open(title: string, status: string): void;
	/** Update the bar (0-100, clamped) and the status line under it. */
	update(progress: number, status: string): void;
	/** Hide the modal. */
	close(): void;
}

export interface ExportProgressModalDeps {
	doc: Document;
	getTranslator(): Translator;
	/**
	 * Where the modal mounts: the viewer's chrome root, so it takes the viewer's theme variables.
	 * Defaults to `doc.body`.
	 */
	getHost?(): HTMLElement;
	/** The Cancel control was pressed (abort the in-flight export). */
	onCancel(): void;
}

/** Build a (closed) progress modal; `open` mounts it into the host. */
export function createExportProgressModal(deps: ExportProgressModalDeps): ExportProgressModal {
	let el: PptxUiExportProgressElement | null = null;
	let title = '';
	let opened = false;

	const show = (open: boolean, progress: number, status: string): void => {
		if (!el) {
			registerPptxWebControls();
			el = deps.doc.createElement('pptx-ui-export-progress');
			el.addEventListener('export-cancel', () => deps.onCancel());
		}
		// The chrome is rebuilt on a locale switch, so the host can change between exports.
		const host = deps.getHost?.() ?? deps.doc.body;
		if (open && el.parentElement !== host) {
			host.appendChild(el);
		}
		el.state = { open, title, progress, status, translate: deps.getTranslator() };
	};

	return {
		open(nextTitle, status) {
			title = nextTitle;
			opened = true;
			show(true, 0, status);
		},
		update(progress, status) {
			// A late progress report after Cancel must not bring the modal back.
			if (opened) {
				show(true, progress, status);
			}
		},
		close() {
			// Kept in place while closed, so it hands focus back to where the export started.
			if (opened) {
				opened = false;
				show(false, 0, '');
			}
		},
	};
}
