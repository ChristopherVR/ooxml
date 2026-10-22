/**
 * viewer-file-io.service.ts: Viewer-scoped state + logic for getting `.pptx`
 * bytes in and out of the viewer: the built-in File ▸ Open override signal (a
 * native-picker-loaded deck swapped in place ahead of the host `content`
 * input), serialising the current deck to bytes, triggering a browser
 * download, and resolving the loaded source bytes for collaboration's
 * elected-writer write-back.
 *
 * Extracted from {@link PowerPointViewerComponent}: the component keeps the
 * public {@link PowerPointViewerComponent.getContent} imperative-handle method
 * (documented host API, so it cannot move off the component) as a thin
 * delegate to {@link getContent} here. The component binds the few accessors
 * it alone owns (canEdit / host `content` input / onOpenFile override / editor
 * slides + template elements / the `contentChange` emitter) via {@link bind}.
 *
 * Provide it once on the viewer component (`providers: [ViewerFileIOService]`).
 */

import { inject, Injectable, signal } from '@angular/core';
import type { DeckSaveIntent, DocumentSaveReason, DocumentStorage } from 'ooxml-ui/pptx';
import {
	deliverSavedPresentation,
	exportDeckJson,
	openPptxFile,
	savedPresentationFileName,
} from 'ooxml-ui/pptx';
import type { PptxSaveFormat, PptxSection, PptxSlide } from 'pptx-viewer-core';

import { ExportService } from './export.service';
import { LoadContentService } from './load-content.service';
import { buildSaveSlides } from './template-mode';
import type { TemplateElementsBySlideId } from './template-mode';

/** Live host accessors the file-IO controller needs. */
interface FileIOHost {
	readonly canEdit: () => boolean;
	readonly content: () => Uint8Array | ArrayBuffer | null;
	readonly onOpenFile: () => (() => void) | undefined;
	readonly slides: () => readonly PptxSlide[];
	readonly sections: () => readonly PptxSection[];
	readonly templateElementsBySlideId: () => TemplateElementsBySlideId;
	readonly emitContentChange: (bytes: Uint8Array) => void;
	/**
	 * File > Info > Protect Presentation state. When it carries a password every
	 * save below serialises through `saveEncrypted`, so the downloaded file is an
	 * encrypted OLE2 container instead of a plain ZIP.
	 */
	readonly saveIntent: () => DeckSaveIntent;
	/**
	 * Options > Accessibility > "feedback with sound", and Options > Save >
	 * "keep the last AutoRecover version": run once a Save/Save-As download
	 * actually completes.
	 */
	readonly afterSuccessfulSave: (format: PptxSaveFormat) => void;
	/** Host `storage` input: when set, Save / Save As go to `storage.save()`, not a download. */
	readonly storage?: () => DocumentStorage | undefined;
	/** Host `fileName` input: the stem the saved copy is named after. */
	readonly fileName?: () => string | undefined;
	/** A Save / Save As failed (serialising, or the host's storage rejected). */
	readonly onSaveError?: (error: unknown) => void;
}

@Injectable()
export class ViewerFileIOService {
	private readonly loader = inject(LoadContentService);
	private readonly exportSvc = inject(ExportService);

	/**
	 * Built-in File ▸ Open override of the host `content` input. The native
	 * picker sets this to swap the deck in place; a fresh `content` input clears
	 * it so external reloads always win.
	 */
	readonly contentOverride = signal<Uint8Array | ArrayBuffer | null>(null);

	private host: FileIOHost | null = null;

	/** Wire the host accessors (called once from the component constructor). */
	bind(host: FileIOHost): void {
		this.host = host;
	}

	private requireHost(): FileIOHost {
		if (!this.host) {
			throw new Error('ViewerFileIOService.bind() was not called');
		}
		return this.host;
	}

	/** The currently active content: a picked-file override, else the host `content` input. */
	activeContent(): Uint8Array | ArrayBuffer | null {
		return this.contentOverride() ?? this.requireHost().content();
	}

	/** The loaded source `.pptx` bytes (for elected-writer write-back), if any. */
	sourceBytes(): Uint8Array | null {
		const content = this.activeContent();
		if (!content) {
			return null;
		}
		return content instanceof Uint8Array ? content : new Uint8Array(content);
	}

	/**
	 * Serialise the current presentation to `.pptx` bytes. When editing, this
	 * serialises the editor's edited deck so changes persist.
	 */
	async getContent(): Promise<Uint8Array> {
		const host = this.requireHost();
		const intent = host.saveIntent();
		const data = host.canEdit()
			? await this.loader.saveSlides(
					buildSaveSlides(host.slides(), host.templateElementsBySlideId()),
					'pptx',
					host.sections(),
					intent,
				)
			: await this.loader.getContent(intent);
		// Mirror React's imperative handle: serialising the deck also notifies the
		// host so listeners wired to (contentChange) receive the latest bytes.
		host.emitContentChange(data);
		return data;
	}

	/**
	 * Serialise the current deck and trigger a browser download of the `.pptx`.
	 * Surfaced on the mobile toolbar so saving is reachable without the desktop
	 * ribbon's File tab.
	 */
	/** The host `storage` input, for chrome that lists recovery copies (Recent, Account). */
	storage(): DocumentStorage | undefined {
		return this.host?.storage?.();
	}

	/** File > Save: `.pptx`, reported to host storage as `save`. */
	async save(): Promise<void> {
		await this.saveAs('pptx', 'save');
	}

	/**
	 * Write `format` (`reason` is the command that ran: Save As unless given). A failure is
	 * reported through `onSaveError` and never rejects, so template handlers cannot leak it.
	 */
	async saveAs(format: PptxSaveFormat, reason: DocumentSaveReason = 'saveAs'): Promise<void> {
		try {
			await this.writeSave(format, reason);
		} catch (error) {
			console.error(`[PowerPointViewer] Save as .${format} failed:`, error);
			this.requireHost().onSaveError?.(error);
		}
	}

	private async writeSave(format: PptxSaveFormat, reason: DocumentSaveReason): Promise<void> {
		const host = this.requireHost();
		const intent = host.saveIntent();
		const slides = buildSaveSlides(host.slides(), host.templateElementsBySlideId());
		const bytes = host.canEdit()
			? await this.loader.saveSlides(slides, format, host.sections(), intent)
			: await this.loader.saveSlides(this.loader.slides(), format, undefined, intent);
		host.emitContentChange(bytes);
		const fileName = savedPresentationFileName(host.fileName?.(), format);
		const storage = host.storage?.();
		if (storage) {
			// The host keeps the file; a rejection propagates and skips the recovery cleanup.
			await deliverSavedPresentation({ storage, bytes, format, fileName, reason });
		} else {
			this.exportSvc.savePresentation(bytes, fileName, format);
		}
		host.afterSuccessfulSave(format);
	}

	async saveAsPptx(): Promise<void> {
		await this.saveAs('pptx');
	}

	async saveAsPpsx(): Promise<void> {
		await this.saveAs('ppsx');
	}

	async saveAsPptm(): Promise<void> {
		await this.saveAs('pptm');
	}

	/** Legacy binary PowerPoint 97-2003 `.ppt`; `saveAs` picks the OLE2 MIME type. */
	async saveAsPpt(): Promise<void> {
		await this.saveAs('ppt');
	}

	/**
	 * File > Export > Export as JSON: serialise the live deck (templates merged
	 * back in when editing) to `pptx-viewer-json` and trigger the download.
	 * `sourceName` (the host `fileName` input) seeds the download filename.
	 */
	exportJson(sourceName?: string | null): void {
		const host = this.requireHost();
		const data = this.loader.parsedData();
		if (!data) {
			return;
		}
		const slides = host.canEdit()
			? buildSaveSlides(host.slides(), host.templateElementsBySlideId())
			: this.loader.slides();
		// `parsedData()` is the deck AS LOADED, so spreading it alone re-exports
		// the file's original show settings and silently drops every edit the
		// Slide Show tab and the Set Up dialog have made since. The save path
		// already reads the live signal; the JSON export has to as well, or
		// "export the live deck" is only true of the slides.
		exportDeckJson(
			{
				...data,
				slides: [...slides],
				presentationProperties: this.loader.presentationProperties(),
			},
			sourceName,
		);
	}

	/**
	 * File ▸ Open: host override (`onOpenFile` input) takes precedence; otherwise
	 * a built-in native picker loads the chosen presentation in place.
	 */
	openFile(): void {
		const host = this.requireHost();
		const override = host.onOpenFile();
		if (override) {
			override();
			return;
		}
		void (async () => {
			const picked = await openPptxFile();
			if (picked) {
				this.contentOverride.set(new Uint8Array(picked.buffer));
			}
		})();
	}
}
