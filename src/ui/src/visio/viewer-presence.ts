import type { VisioPresence } from 'ooxml-core/visio/ui';
import type { CollabSession } from 'ooxml-core/collab';
import type { ViewerController, ViewerState } from './controller';
import { type PeerPresence, presenceScale, renderPresenceOverlay } from './viewer-presence-overlay';
import { pagePoint } from './viewer-page-point';

/**
 * Live cursors and selections of one shared drawing: publishes this window's page, pointer and
 * selection through the session's presence (already throttled there) and draws the peers that
 * are on the same page over the canvas. It borrows the session; the caller owns its lifetime.
 */
export class ViewerPresence {
	readonly #stops: (() => void)[] = [];
	#peers: PeerPresence[] = [];
	#queued = false;
	#disposed = false;
	constructor(
		private readonly session: CollabSession<VisioPresence>,
		private readonly controller: ViewerController,
		private readonly viewport: HTMLElement | null,
	) {
		this.#stops.push(
			session.on('peers', (peers) => {
				this.#peers = peers.filter((peer) => typeof peer.page === 'string');
				this.#schedule();
			}),
			controller.subscribe((state) => {
				this.#publishLocation(state);
				this.#schedule();
			}),
		);
		this.#peers = session.peers().filter((peer) => typeof peer.page === 'string');
		this.#publishLocation(controller.state);
		this.#schedule();
		if (viewport) {
			const move = (event: PointerEvent): void => this.#point(event);
			const leave = (): void => this.session.updatePresence({ x: undefined, y: undefined });
			viewport.addEventListener('pointermove', move);
			viewport.addEventListener('pointerleave', leave);
			this.#stops.push(
				() => viewport.removeEventListener('pointermove', move),
				() => viewport.removeEventListener('pointerleave', leave),
			);
		}
	}
	/** The remote collaborators currently drawn (same page or not). */
	get peers(): readonly PeerPresence[] {
		return this.#peers;
	}
	dispose(): void {
		this.#disposed = true;
		for (const stop of this.#stops.splice(0)) stop();
		this.#peers = [];
		this.#paint();
		// Leave no stale location behind in a session that outlives this viewer.
		this.session.updatePresence({ x: undefined, y: undefined, selection: undefined });
	}
	#publishLocation(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		const pageId = page?.id ?? '';
		const selection = page
			? state.selectedShapes
					.filter((shape) => (shape.pageId ?? pageId) === pageId)
					.map((shape) => shape.id)
			: [];
		this.session.updatePresence({
			page: pageId,
			selection: selection.length ? selection.slice(0, 64) : undefined,
		});
	}
	#point(event: PointerEvent): void {
		const { document, pageIndex } = this.controller.state;
		const page = document?.pages[pageIndex];
		const svg = this.viewport?.querySelector<SVGSVGElement>('svg');
		const point =
			page && svg ? pagePoint(svg, page, event, { snap: false, bounded: false }) : undefined;
		const inside =
			point &&
			page &&
			point.x >= 0 &&
			point.y >= 0 &&
			point.x <= page.width &&
			point.y <= page.height;
		this.session.updatePresence(
			inside ? { x: point.x, y: point.y } : { x: undefined, y: undefined },
		);
	}
	/** Paints after the viewer's own render of the same state, which replaces the page SVG. */
	#schedule(): void {
		if (this.#queued) return;
		this.#queued = true;
		queueMicrotask(() => {
			this.#queued = false;
			if (!this.#disposed) this.#paint();
		});
	}
	#paint(): void {
		const svg = this.viewport?.querySelector<SVGSVGElement>('svg');
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!svg) return;
		if (!page) {
			svg.querySelector(':scope > [data-presence-overlay]')?.remove();
			return;
		}
		renderPresenceOverlay(svg, page, this.#peers, presenceScale(svg, state.zoom));
	}
}
