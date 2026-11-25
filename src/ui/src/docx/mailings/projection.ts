import type { DocumentModel } from 'ooxml-core/docx';
import type { EditorCore } from '../editor-core';
import { createPrintLayoutController, type PrintLayoutController } from '../print-layout-view';
import { emit } from '../events';
import { mailingsText } from './strings';

/** The paginated preview and print share the template's media, never its editing view. */
export class MailingsProjection {
	private layout: PrintLayoutController | undefined;
	private preview: DocumentModel | undefined;
	private scroll: { top: number; left: number } | undefined;
	private status: HTMLOutputElement | undefined;

	constructor(private readonly core: EditorCore) {}

	show(model: DocumentModel, record: number, count: number, skipped: boolean): void {
		this.preview = model;
		this.mount()?.relayoutNow(model);
		if (this.status) {
			const { locale } = this.core;
			const label = `${mailingsText(locale, 'Record')} ${record} / ${count}`;
			this.status.value = skipped
				? `${label} · ${mailingsText(locale, 'This recipient is skipped by a merge rule.')}`
				: label;
		}
	}

	stop(): void {
		this.preview = undefined;
		this.unmount();
	}

	print(model: DocumentModel): void {
		const layout = this.mount();
		try {
			layout?.print(model, (text) => emit(this.core.element, 'document-warning', text));
		} finally {
			if (this.preview) layout?.relayoutNow(this.preview);
			else this.unmount();
		}
	}

	private mount(): PrintLayoutController | undefined {
		if (this.layout) return this.layout;
		const { core } = this;
		const { canvas, paper } = core.shell;
		if (!canvas) return undefined;
		this.scroll = { top: canvas.scrollTop, left: canvas.scrollLeft };
		this.layout = createPrintLayoutController(
			canvas,
			() => {},
			(part, contentType) => core.imageMedia.urlFor(part, contentType),
			undefined,
			() => core.reviewDisplayMode,
		);
		this.layout.element.dataset.mergeProjection = '';
		this.layout.setActive(true);
		this.status = document.createElement('output');
		this.status.className = 'dve-mailings-record';
		canvas.append(this.status, this.layout.element);
		canvas.dataset.mailMergePreview = '';
		if (paper) paper.inert = true;
		canvas.scrollTop = 0;
		return this.layout;
	}

	private unmount(): void {
		this.layout?.element.remove();
		this.status?.remove();
		this.status = undefined;
		this.layout = undefined;
		const { canvas, paper } = this.core.shell;
		if (!canvas?.hasAttribute('data-mail-merge-preview')) return;
		delete canvas.dataset.mailMergePreview;
		if (paper) paper.inert = false;
		if (this.scroll) {
			canvas.scrollTop = this.scroll.top;
			canvas.scrollLeft = this.scroll.left;
		}
		this.scroll = undefined;
	}
}
