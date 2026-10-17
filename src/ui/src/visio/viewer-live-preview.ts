import type { VisioDocument } from 'ooxml-core/visio';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController } from './controller';

/** Hovering a tile must rest this long before its preview is worked out. */
const SETTLE_MS = 120;

type GalleryDetail = { gallery?: string; itemId?: string };

/**
 * Live Preview: while the pointer or keyboard rests on a tile of a gallery, the drawing shows the
 * result of picking it, and goes back when the tile is left. The preview is made by firing the
 * same pick the tile would, inside `controller.previewing`, so each gallery's own handler decides
 * what the tile does and nothing is applied: the preview cannot differ from the real thing, and
 * history is never touched. Galleries opt in by setting `preview` on their state.
 */
export class ViewerLivePreview {
	enabled = true;
	#timer: ReturnType<typeof setTimeout> | undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		/** Draw `document` in place of the drawing. */
		private readonly show: (document: VisioDocument) => void,
		/** Go back to the real drawing. */
		private readonly end: () => void,
	) {}
	wire(): () => void {
		const Abort = this.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.root.addEventListener(
			'office-gallery-preview',
			(event) => this.#hover(event as CustomEvent<GalleryDetail>),
			options,
		);
		this.root.addEventListener('office-gallery-preview-end', () => this.stop(), options);
		return () => {
			events.abort();
			this.stop();
		};
	}
	/** Forget any pending or shown preview. */
	stop(): void {
		clearTimeout(this.#timer);
		this.#timer = undefined;
		this.controller.cancelPreview();
		this.end();
	}
	#hover(event: CustomEvent<GalleryDetail>): void {
		if (!this.enabled) return;
		const gallery = event
			.composedPath()
			.find(
				(node): node is OfficeUiGallery =>
					node instanceof HTMLElement && node.localName === 'office-ui-gallery',
			);
		const detail = event.detail;
		if (!gallery || !detail?.itemId) return;
		clearTimeout(this.#timer);
		this.#timer = setTimeout(() => this.#preview(gallery, detail), SETTLE_MS);
	}
	#preview(gallery: OfficeUiGallery, detail: GalleryDetail): void {
		this.#timer = undefined;
		this.controller.previewing(
			(document) => {
				// A refused, stale or cancelled preview leaves what is shown as it is.
				if (document && this.#timer === undefined) this.show(document);
			},
			// Not bubbling: only the gallery's own pick handler sees it.
			() =>
				gallery.dispatchEvent(
					new CustomEvent('office-gallery-pick', { detail, bubbles: false, composed: false }),
				),
		);
	}
}
