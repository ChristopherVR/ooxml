/**
 * Word's webfont fallback: the `googleFonts` option of `<docx-editor>`. Off unless the host sets
 * it. When on, the families a document draws that the reader has neither installed nor embedded
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
	prepareWebfonts,
	syncGoogleWebfontStylesheet,
	type PrepareWebfontsOptions,
	type WebfontReadiness,
} from '../fonts/webfonts';

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
	option: GoogleFontsOption | undefined;
	readonly linkId = `${DOCX_WEBFONT_LINK_PREFIX}-${++instances}`;

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
		const base = docxGoogleFontsBase(this.option);
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

	/** Whether this editor's `<link>` is in `doc`. */
	linked(doc: Document | undefined): boolean {
		return Boolean(doc?.getElementById(this.linkId));
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
