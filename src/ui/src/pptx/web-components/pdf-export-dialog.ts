import { defineElement } from '../../registry';
import {
	initialPdfExportChoice,
	resolvePdfExportSelection,
	validatePdfExportChoice,
} from '../render/pdf-export-range';
import type { PdfExportChoice } from '../render/pdf-export-range';
import type { PdfSlideSelection } from '../export/pdf-vector-select';
import { createInsertDialogFrame } from './insert-dialog-frame';
import type { InsertDialogFrame } from './insert-dialog-frame';
import { createPdfExportDialogBody } from './pdf-export-dialog-body';
import type { PdfExportDialogBody } from './pdf-export-dialog-body';

/** Translates a dictionary key with `{{name}}` params; an untranslated key comes back unchanged. */
export type PdfExportDialogTranslate = (
	key: string,
	params?: Record<string, string | number>,
) => string;

/** What Export hands the host: pass `selection` to the vector PDF export. */
export interface PdfExportDialogDetail {
	choice: PdfExportChoice;
	selection: PdfSlideSelection;
}
export type PdfExportDialogConfirmEvent = CustomEvent<PdfExportDialogDetail>;

export interface PptxUiPdfExportDialogElement extends HTMLElement {
	open: boolean;
	/** The deck's slides (only `hidden` is read), for the range check. */
	slides: readonly { hidden?: boolean }[];
	/** 0-based index of the current slide. */
	currentSlide: number;
	/** "Include hidden slides" when the dialog opens (Options > Advanced > Save as PDF). */
	includeHiddenSlides: boolean;
	translateLabel: PdfExportDialogTranslate | undefined;
}

const FALLBACK: Record<string, string> = {
	'pptx.pdfExport.dialog.title': 'PDF Options',
	'pptx.pdfExport.dialog.range': 'Range',
	'pptx.pdfExport.dialog.all': 'All',
	'pptx.pdfExport.dialog.current': 'Current slide',
	'pptx.pdfExport.dialog.slides': 'Slides',
	'pptx.pdfExport.dialog.from': 'From:',
	'pptx.pdfExport.dialog.to': 'To:',
	'pptx.pdfExport.dialog.includeHidden': 'Include hidden slides',
	'pptx.pdfExport.dialog.export': 'Export',
	'pptx.pdfExport.dialog.rangeInvalid': 'Enter slide numbers from 1 to {{count}}.',
	'pptx.pdfExport.dialog.rangeEmpty': 'There are no slides to export in this range.',
	'pptx.common.cancel': 'Cancel',
	'pptx.smartart.close': 'Close',
};

/**
 * `pptx-ui-pdf-export-dialog`: the options step of File > Export > PDF, like PowerPoint's Save as
 * PDF options: Range (All, Current slide, Slides from-to) and "Include hidden slides", in the
 * shared `office-ui-dialog` with Cancel and Export. Controlled: the host sets `open` (with
 * `slides`, `currentSlide`, `includeHiddenSlides` and `translateLabel`); Export emits
 * `pdf-export-confirm` with {@link PdfExportDialogDetail} once the range is valid, and every
 * close (Export, Cancel, Escape, the backdrop) emits `dialog-close` so the host clears its open
 * state. The host runs the export. Range logic: `render/pdf-export-range.ts`.
 */
export function definePptxPdfExportDialog(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-pdf-export-dialog')) {
		return;
	}

	class PdfExportDialog extends HTMLElement implements PptxUiPdfExportDialogElement {
		static observedAttributes = ['open'];
		#frame: InsertDialogFrame | undefined;
		#body: PdfExportDialogBody | undefined;
		#translate: PdfExportDialogTranslate | undefined;
		#slides: readonly { hidden?: boolean }[] = [];
		#current = 0;
		#includeHidden = false;
		/** The user changed a control since the dialog opened; until then new inputs re-seed it. */
		#touched = false;
		#seeding = false;

		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		get slides(): readonly { hidden?: boolean }[] {
			return this.#slides;
		}
		set slides(value: readonly { hidden?: boolean }[]) {
			this.#slides = value ?? [];
			this.#reseed();
		}
		get currentSlide(): number {
			return this.#current;
		}
		set currentSlide(value: number) {
			this.#current = Number.isFinite(value) ? Math.max(0, Math.floor(value)) : 0;
			this.#reseed();
		}
		get includeHiddenSlides(): boolean {
			return this.#includeHidden;
		}
		set includeHiddenSlides(value: boolean) {
			this.#includeHidden = Boolean(value);
			this.#reseed();
		}
		get translateLabel(): PdfExportDialogTranslate | undefined {
			return this.#translate;
		}
		set translateLabel(value: PdfExportDialogTranslate | undefined) {
			if (value === this.#translate) {
				return;
			}
			this.#translate = value;
			this.#sync();
		}

		connectedCallback(): void {
			this.#build();
			this.#sync();
			this.#show();
		}

		attributeChangedCallback(): void {
			if (this.isConnected) {
				this.#show();
			}
		}

		#text(key: string, params?: Record<string, string | number>): string {
			const value = this.#translate?.(key, params);
			if (value && value !== key) {
				return value;
			}
			const fallback = FALLBACK[key] ?? key;
			return fallback.replace(/\{\{(\w+)\}\}/gu, (_, name: string) => String(params?.[name] ?? ''));
		}

		#build(): void {
			if (this.#frame) {
				return;
			}
			const body = createPdfExportDialogBody(
				this.ownerDocument,
				() => {
					if (!this.#seeding) this.#touched = true;
					this.#validate();
				},
				() => this.#export(),
			);
			this.#body = body;
			this.#frame = createInsertDialogFrame(
				this,
				body.element,
				() => this.#export(),
				() => this.#closed(),
			);
			this.#frame.dialog.style.setProperty(
				'--office-dialog-width',
				'min(420px, calc(100vw - 32px))',
			);
			this.#frame.dialog.dataset.testid = 'pptx-pdf-export-dialog';
		}

		#sync(): void {
			if (!this.#frame || !this.#body) {
				return;
			}
			this.#body.label({
				range: this.#text('pptx.pdfExport.dialog.range'),
				all: this.#text('pptx.pdfExport.dialog.all'),
				current: this.#text('pptx.pdfExport.dialog.current'),
				slides: this.#text('pptx.pdfExport.dialog.slides'),
				from: this.#text('pptx.pdfExport.dialog.from'),
				to: this.#text('pptx.pdfExport.dialog.to'),
				includeHidden: this.#text('pptx.pdfExport.dialog.includeHidden'),
			});
			this.#validate();
		}

		/** Show the error of the shown choice (Export is disabled while there is one) and return it. */
		#validate(): string | undefined {
			const body = this.#body;
			const frame = this.#frame;
			if (!body || !frame) {
				return undefined;
			}
			const error = validatePdfExportChoice(body.read(), this.#slides, this.#current);
			const message = error
				? this.#text(`pptx.pdfExport.dialog.${error}`, { count: this.#slides.length })
				: undefined;
			body.error(message ?? '');
			frame.sync(
				this.#text('pptx.pdfExport.dialog.title'),
				this.#text('pptx.smartart.close'),
				this.#text('pptx.common.cancel'),
				this.#text('pptx.pdfExport.dialog.export'),
				!error,
			);
			return message;
		}

		#show(): void {
			const frame = this.#frame;
			const body = this.#body;
			if (!frame || !body || frame.dialog.open === this.open) {
				return;
			}
			if (this.open) {
				this.#touched = false;
				this.#seed();
			}
			frame.dialog.open = this.open;
		}

		/** Show the opening choice for the current inputs. */
		#seed(): void {
			if (!this.#body) return;
			this.#seeding = true;
			try {
				this.#body.set(
					initialPdfExportChoice(this.#slides.length, this.#includeHidden),
					this.#slides.length,
				);
			} finally {
				this.#seeding = false;
			}
			this.#sync();
		}

		/** Inputs set after `open` (in any order) still decide the opening choice. */
		#reseed(): void {
			if (this.open && this.#frame?.dialog.open && !this.#touched) this.#seed();
			else this.#validate();
		}

		#export(): void {
			const body = this.#body;
			if (!body) {
				return;
			}
			if (this.#validate() !== undefined) {
				return;
			}
			const choice = body.read();
			this.dispatchEvent(
				new CustomEvent<PdfExportDialogDetail>('pdf-export-confirm', {
					detail: { choice, selection: resolvePdfExportSelection(choice, this.#current) },
					bubbles: true,
					composed: true,
				}),
			);
			if (this.#frame) {
				this.#frame.dialog.open = false;
			}
			this.#closed();
		}

		#closed(): void {
			this.removeAttribute('open');
			this.dispatchEvent(new CustomEvent('dialog-close', { bubbles: true, composed: true }));
		}
	}

	defineElement(registry, 'pptx-ui-pdf-export-dialog', PdfExportDialog);
}
