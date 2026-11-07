import { collectVisioWebfontRuns, type VisioDocument } from 'ooxml-core/visio';
import { resolveGoogleFontsCss2Base, type GoogleFontsOption } from 'ooxml-core/text';
import { FONT_DOWNLOADS_DEFAULT, effectiveFontDownloads } from '../fonts/font-download-policy';
import {
	prefetchWebfonts,
	prepareWebfonts,
	syncGoogleWebfontStylesheet,
	type WebfontReadiness,
} from '../fonts/webfonts';

/** The managed `<link>` id of the first viewer in a document; later ones add a number. */
export const VISIO_WEBFONT_LINK_ID = 'visio-viewer-webfonts';
let viewers = 0;

/**
 * The `googleFonts` webfont fallback of one `<visio-viewer>`: which CSS2 base the host's option
 * and the reader's Trust Center choice resolve to (both unset means off, unlike PowerPoint), and
 * the managed `<link>` in the host
 * document's head (a shadow root's `@font-face` rules would not apply). Each viewer owns its own
 * link, so two viewers on one page never remove each other's faces.
 */
export class ViewerWebfonts {
	#option: GoogleFontsOption | undefined;
	#readerAllows = FONT_DOWNLOADS_DEFAULT;
	readonly linkId: string;
	/** Bumped by every request and option change, so a superseded wait is ignored. */
	#generation = 0;
	constructor(
		private readonly host: () => Document,
		private readonly prepare: typeof prepareWebfonts = prepareWebfonts,
	) {
		const index = ++viewers;
		this.linkId = index === 1 ? VISIO_WEBFONT_LINK_ID : `${VISIO_WEBFONT_LINK_ID}-${index}`;
	}
	get option(): GoogleFontsOption | undefined {
		return this.#option;
	}
	set option(value: GoogleFontsOption | undefined) {
		this.#option = value;
		this.#changed();
	}
	/** The reader's Trust Center choice, used only while the host left `option` unset. */
	get readerAllows(): boolean {
		return this.#readerAllows;
	}
	set readerAllows(value: boolean) {
		this.#readerAllows = value;
		this.#changed();
	}
	#changed(): void {
		++this.#generation;
		// Off: drop the stylesheet now, so nothing more is requested from the mirror or Google.
		if (this.base() === null) this.remove();
	}
	/**
	 * The CSS2 base requested, or `null` when the fallback is off: the host's option when set,
	 * else Google Fonts when the reader allowed downloads (`effectiveFontDownloads`).
	 */
	base(): string | null {
		return resolveGoogleFontsCss2Base(
			effectiveFontDownloads(this.#option, this.#readerAllows),
			null,
			'visio-viewer',
		);
	}
	/**
	 * Load the faces the drawing draws in families the reader has neither installed nor gets from
	 * the bundled catalogue under another name. Resolves `undefined` when off, superseded or when
	 * no webfont was needed; otherwise the readiness, after the bounded wait. Never rejects.
	 */
	async load(document: VisioDocument | null): Promise<WebfontReadiness | undefined> {
		const generation = ++this.#generation;
		const base = this.base();
		if (base === null || !document) return undefined;
		let readiness: WebfontReadiness;
		try {
			readiness = await this.prepare({
				document: this.host(),
				linkId: this.linkId,
				runs: collectVisioWebfontRuns(document),
				base,
			});
		} catch {
			return undefined;
		}
		if (generation !== this.#generation || readiness.families.length === 0) return undefined;
		return readiness;
	}
	/**
	 * Start downloading the faces the file declares before it is parsed (see `prefetchWebfonts`),
	 * so they are usually ready when {@link load} waits. Never throws.
	 */
	prefetch(bytes: Uint8Array | ArrayBuffer): void {
		try {
			prefetchWebfonts({ document: this.host(), linkId: this.linkId, bytes, base: this.base() });
		} catch {
			// A detached or closed document has no head to add to.
		}
	}
	/** Remove this viewer's stylesheet (option off, viewer destroyed). */
	remove(): void {
		++this.#generation;
		try {
			syncGoogleWebfontStylesheet(this.host(), this.linkId, null);
		} catch {
			// A detached or closed document has no head to clean.
		}
	}
}
