/**
 * Word's webfont fallback: the `googleFonts` option of `<docx-editor>` when the host sets it,
 * otherwise the reader's Trust Center choice (`font-downloads.ts`, off until turned on), joined by
 * the shared `effectiveFontDownloads` rule. When on, the families a document draws that the reader has neither installed nor embedded
 * (Word does not load a document's embedded fonts yet, so none count as embedded) are requested
 * from the Google Fonts CSS2 API, or the host's mirror, through one managed `<link>` per editor;
 * a family the catalogue does not serve under its own name ("Calibri") is requested as its
 * metric-compatible clone ("Carlito"), which Word's font stacks (`cssFontStack`) already list
 * after the document font in both measurement and paint.
 *
 * Which runs and families is `collectDocxWebfontRuns` (core); the browser side is
 * `fonts/webfonts`. This module only joins them to the element's option.
 *
 * @module docx/docx-webfonts
 */
import type { DocumentModel } from 'ooxml-core/docx';
import { collectDocxWebfontRuns } from 'ooxml-core/docx/layout';
import {
	resolveGoogleFontsCss2Base,
	type GoogleFontsOption,
	type WebfontRun,
} from 'ooxml-core/text';
import {
	prefetchLinkId,
	prefetchWebfonts,
	prepareWebfonts,
	syncGoogleWebfontStylesheet,
	type PrepareWebfontsOptions,
	type WebfontReadiness,
} from '../fonts/webfonts';
import { effectiveFontDownloads, fontDownloadLock } from '../fonts/font-download-policy';
import { readerAllowsFontDownloads, setReaderAllowsFontDownloads } from './font-downloads';

/** Prefix of each editor's managed `<link>` id (`docx-editor-webfonts-1`, ...). */
export const DOCX_WEBFONT_LINK_PREFIX = 'docx-editor-webfonts';

const OFF: WebfontReadiness = { href: null, families: [], unavailable: [], timedOut: false };

/** The CSS2 base `option` requests, or `null` (off). Unset is off in Word. */
export function docxGoogleFontsBase(option: GoogleFontsOption | undefined): string | null {
	return resolveGoogleFontsCss2Base(option, null, 'docx-editor');
}

/** The document's runs, or none when the model cannot be read (the fallback never throws). */
function runsOf(model: DocumentModel): WebfontRun[] {
	try {
		return collectDocxWebfontRuns(model);
	} catch {
		return [];
	}
}

export type DocxWebfontWaitOptions = Pick<
	PrepareWebfontsOptions,
	'timeoutMs' | 'progressDelayMs' | 'onProgress' | 'isLocallyInstalled'
>;

let instances = 0;

/**
 * One editor's webfont state: the host's option and its own `<link>`, so two editors on a page
 * (the suite's tabs) never replace each other's stylesheet.
 */
export class DocxWebfonts {
	/** The host's `googleFonts` value; unset leaves the choice to the reader. */
	option: GoogleFontsOption | undefined;
	readonly linkId = `${DOCX_WEBFONT_LINK_PREFIX}-${++instances}`;

	/** The CSS2 base the loader uses: the host's value, else the reader's Trust Center choice. */
	base(): string | null {
		return docxGoogleFontsBase(effectiveFontDownloads(this.option, readerAllowsFontDownloads()));
	}

	/**
	 * Point this editor's `<link>` at the faces `model` needs and wait (bounded) for them; with
	 * the option off the `<link>` is removed and nothing is requested. Never rejects.
	 */
	async prepare(
		doc: Document | undefined,
		model: DocumentModel,
		wait: DocxWebfontWaitOptions = {},
	): Promise<WebfontReadiness> {
		if (!doc?.head) return OFF;
		const base = this.base();
		if (base === null) {
			syncGoogleWebfontStylesheet(doc, this.linkId, null);
			return OFF;
		}
		try {
			return await prepareWebfonts({
				document: doc,
				linkId: this.linkId,
				runs: runsOf(model),
				embedded: [],
				base,
				...wait,
			});
		} catch {
			return OFF;
		}
	}

	/**
	 * Start downloading the faces `bytes` declares before it is parsed (see `prefetchWebfonts`),
	 * so they are usually ready when {@link prepare} waits. Never throws.
	 */
	prefetch(doc: Document | undefined, bytes: Uint8Array | ArrayBuffer): void {
		if (!doc?.head) return;
		prefetchWebfonts({
			document: doc,
			linkId: this.linkId,
			bytes,
			base: this.base(),
		});
	}

	/** Whether this editor's `<link>` (or its early `-prefetch` one) is in `doc`. */
	linked(doc: Document | undefined): boolean {
		return Boolean(
			doc?.getElementById(this.linkId) ?? doc?.getElementById(prefetchLinkId(this.linkId)),
		);
	}
}

/** What {@link refreshDocxWebfonts} needs from the editor. */
export interface DocxWebfontHost {
	readonly element: HTMLElement;
	readonly webfonts: DocxWebfonts;
	readonly model: DocumentModel;
	readonly shell: { printLayout?: { remeasure(): void } };
}

/**
 * Bring the webfonts up to date for the document already on screen (a document set without
 * `load`, or the option changed). Faces that arrive fire `document.fonts` `loadingdone`, on which
 * Print Layout measures again and the run decorations refresh; a stylesheet that was removed
 * takes its faces away without that event, so Print Layout is told to measure again here.
 */
export async function refreshDocxWebfonts(host: DocxWebfontHost): Promise<void> {
	const doc = host.element.ownerDocument;
	const had = host.webfonts.linked(doc);
	await host.webfonts.prepare(doc, host.model);
	if (had && !host.webfonts.linked(doc)) host.shell.printLayout?.remeasure();
}

/** The Trust Center toggle as File > Options shows it: the host's value, locked, or the reader's. */
export function fontDownloadOptions(webfonts: DocxWebfonts): {
	allowFontDownloads: boolean;
	fontDownloadsLocked: boolean;
} {
	const locked = fontDownloadLock(webfonts.option);
	return {
		allowFontDownloads: locked ?? readerAllowsFontDownloads(),
		fontDownloadsLocked: locked !== undefined,
	};
}

/**
 * Records the reader's choice and applies it to the open document at once: on requests its
 * fonts, off removes the stylesheets and measures again. A host's value wins, so nothing changes
 * then.
 */
export function setFontDownloads(host: DocxWebfontHost, allow: boolean): void {
	if (fontDownloadLock(host.webfonts.option) !== undefined) return;
	setReaderAllowsFontDownloads(allow);
	void refreshDocxWebfonts(host);
}
