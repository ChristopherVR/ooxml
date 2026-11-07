/**
 * The webfont fallback of `<xlsx-editor>` (the `googleFonts` property): for workbook fonts the
 * reader has neither installed nor embedded, one managed `<link>` loads the face the bundled
 * Google Fonts catalogue serves (or its verified metric-compatible clone, which every workbook
 * font stack lists right after the font, see `workbookFontFamily`). Who decides is the shared
 * rule (`fonts/font-download-policy`): the host's value when it set one, otherwise the reader's
 * Trust Center choice (`font-downloads.ts`), which is off until the reader turns it on.
 *
 * A file opened with `load()` waits (bounded, see `prepareWebfonts`) before it is shown, while
 * `loading` is true; a workbook assigned directly is shown at once. Either way, once faces arrive
 * (also later, through `document.fonts` `loadingdone` while connected) the grid measures text
 * again and repaints, keeping the selection, the scroll and an edit in progress.
 *
 * @module xlsx/webfonts
 */
import type { Workbook } from 'ooxml-core/xlsx';
import { collectWorkbookWebfontRuns } from 'ooxml-core/xlsx';
import { resolveGoogleFontsCss2Base, type GoogleFontsOption } from 'ooxml-core/text';
import {
	prefetchWebfonts,
	prepareWebfonts,
	syncGoogleWebfontStylesheet,
	type WebfontReadiness,
} from '../fonts/webfonts';
import { effectiveFontDownloads, fontDownloadLock } from '../fonts/font-download-policy';
import { readerAllowsFontDownloads, setReaderAllowsFontDownloads } from './font-downloads';

/** Id of the first editor's managed stylesheet link; later editors add `-2`, `-3`... */
export const XLSX_WEBFONT_LINK_ID = 'xlsx-editor-webfonts';

const OFF: WebfontReadiness = { href: null, families: [], unavailable: [], timedOut: false };

let instances = 0;

interface Prepared {
	base: string;
	promise: Promise<WebfontReadiness>;
	settled: boolean;
}

export interface XlsxWebfontsHost {
	readonly host: HTMLElement;
	/** The workbook on screen. */
	workbook(): Workbook | undefined;
	/** Measures text again and repaints (the grid's `remeasure`). */
	remeasure(): void;
}

export class XlsxWebfonts {
	/** The host's `googleFonts` value, as given. */
	option: GoogleFontsOption | undefined = undefined;
	/** This editor's managed `<link>` id (one per editor, so two editors never fight over it). */
	readonly linkId: string;
	readonly #owner: XlsxWebfontsHost;
	#prepared = new WeakMap<Workbook, Prepared>();
	#fonts: FontFaceSet | undefined;
	#queued = false;
	readonly #onFonts = () => this.schedule();

	constructor(owner: XlsxWebfontsHost) {
		this.#owner = owner;
		instances += 1;
		this.linkId = instances === 1 ? XLSX_WEBFONT_LINK_ID : `${XLSX_WEBFONT_LINK_ID}-${instances}`;
	}

	/** The option the loader uses: the host's value when set, else the reader's choice. */
	effective(): GoogleFontsOption {
		return effectiveFontDownloads(this.option, readerAllowsFontDownloads());
	}

	/** The value the Trust Center toggle shows read-only, or `undefined` when the reader decides. */
	lock(): boolean | undefined {
		return fontDownloadLock(this.option);
	}

	/** Whether the Trust Center toggle is on: the host's locked value, else the reader's choice. */
	allowed(): boolean {
		return this.lock() ?? readerAllowsFontDownloads();
	}

	/** The reader's Trust Center choice: stored, and applied to the open workbook at once. */
	setReaderAllows(allow: boolean): void {
		setReaderAllowsFontDownloads(allow);
		this.apply();
	}

	/** The CSS2 base the option requests, or `null` when the fallback is off (the default). */
	base(): string | null {
		return resolveGoogleFontsCss2Base(this.effective(), null, 'xlsx-editor');
	}

	/**
	 * After the host value or the reader's choice changed: on, prepare the open workbook's faces
	 * and measure again; off, remove the stylesheets and repaint with installed or substitute fonts.
	 */
	apply(): void {
		if (this.base() !== null) {
			this.shown(this.#owner.workbook());
			return;
		}
		this.#off(this.#owner.host.ownerDocument);
		this.schedule();
	}

	/** Removes this editor's links and forgets the waits they served. */
	#off(doc: Document): void {
		syncGoogleWebfontStylesheet(doc, this.linkId, null);
		this.#prepared = new WeakMap();
	}

	/**
	 * Start downloading the faces the file declares before it is parsed (see `prefetchWebfonts`),
	 * so they are usually ready when {@link prepare} waits. Never throws.
	 */
	prefetch(bytes: Uint8Array | ArrayBuffer): void {
		const doc = this.#owner.host.ownerDocument;
		if (!doc?.head) return;
		prefetchWebfonts({ document: doc, linkId: this.linkId, bytes, base: this.base() });
	}

	/**
	 * Injects the stylesheet for the workbook's missing fonts and waits (bounded) for their faces.
	 * Off, it removes this editor's link. Never rejects; one wait per workbook and base.
	 */
	prepare(workbook: Workbook): Promise<WebfontReadiness> {
		const doc = this.#owner.host.ownerDocument;
		const base = this.base();
		if (base === null) {
			this.#off(doc);
			return Promise.resolve(OFF);
		}
		const cached = this.#prepared.get(workbook);
		if (cached?.base === base) return cached.promise;
		let runs: ReturnType<typeof collectWorkbookWebfontRuns>;
		try {
			runs = collectWorkbookWebfontRuns(workbook);
		} catch {
			return Promise.resolve(OFF);
		}
		const entry: Prepared = { base, settled: false, promise: Promise.resolve(OFF) };
		entry.promise = prepareWebfonts({ document: doc, linkId: this.linkId, runs, base })
			.catch(() => OFF)
			.finally(() => {
				entry.settled = true;
			});
		this.#prepared.set(workbook, entry);
		return entry.promise;
	}

	/**
	 * A workbook is on screen (opened, assigned, or the option changed): prepare its faces when
	 * that has not happened yet, and measure again once they arrive.
	 */
	shown(workbook: Workbook | undefined): void {
		if (!workbook) {
			if (this.base() === null) this.#off(this.#owner.host.ownerDocument);
			return;
		}
		const waited = this.#prepared.get(workbook);
		const ready = waited?.settled === true && waited.base === this.base();
		void this.prepare(workbook).then((readiness) => {
			if (!ready && readiness.families.length && this.#owner.workbook() === workbook)
				this.schedule();
		});
	}

	/** Coalesces face arrivals into one remeasure per frame while connected. */
	schedule(): void {
		if (this.#queued || !this.#owner.host.isConnected) return;
		this.#queued = true;
		const view = this.#owner.host.ownerDocument.defaultView;
		const run = () => {
			this.#queued = false;
			if (this.#owner.host.isConnected && this.#owner.workbook()) this.#owner.remeasure();
		};
		if (view?.requestAnimationFrame) view.requestAnimationFrame(run);
		else setTimeout(run, 0);
	}

	/** Listens for faces that arrive after the bounded wait (any font the grid draws). */
	connect(): void {
		const fonts = this.#owner.host.ownerDocument.fonts as FontFaceSet | undefined;
		if (!fonts || typeof fonts.addEventListener !== 'function' || this.#fonts === fonts) return;
		this.disconnect();
		this.#fonts = fonts;
		fonts.addEventListener('loadingdone', this.#onFonts);
	}

	disconnect(): void {
		this.#fonts?.removeEventListener('loadingdone', this.#onFonts);
		this.#fonts = undefined;
	}
}
