import type { CollabSession } from 'ooxml-core/collab';
import type { DrawingBinding, DrawingEditor } from 'ooxml-core/visio/collab';
import type { ViewerController } from './controller';

export type Collab = typeof import('ooxml-core/collab');
export type DrawingCollab = typeof import('ooxml-core/visio/collab');
export type Session = ReturnType<Collab['createCollabSession']>;

export interface DrawingSyncOptions {
	/** `join` never seeds an empty room with the drawing this window already has open. */
	join: boolean;
	/** Reports a failure to share or apply a change. */
	fail(message: string): void;
}

/** The DOM-free editor the shared-drawing binding drives, over a viewer controller. */
export function controllerEditor(controller: ViewerController): DrawingEditor {
	let applying = false;
	/** The drawing last shared or adopted here, so a load or edit is published once. */
	let seen = controller.state.document;
	const busy = () => controller.state.edit.busy || controller.state.loading;
	return {
		exportBytes() {
			const { edit } = controller.state;
			if (!edit.sourceAvailable || busy()) return undefined;
			try {
				return controller.exportVsdx().bytes;
			} catch {
				return undefined;
			}
		},
		async applyRemote(bytes, change) {
			const { document, edit } = controller.state;
			applying = true;
			try {
				if (document && edit.sourceAvailable) await controller.applyRemoteSource(bytes, change);
				else await controller.load(bytes);
			} finally {
				applying = false;
				seen = controller.state.document;
			}
		},
		onLocalChange(listener) {
			// The controller drops an event for later listeners when an earlier one changes state (a
			// host fitting the page on load), so a drawing the room has not seen is also found from
			// state: sharing must not depend on listener order.
			const check = () => {
				const { document } = controller.state;
				if (applying || busy() || !document || document === seen) return;
				seen = document;
				listener();
			};
			const stops = [
				controller.onEvent((name, detail) => {
					if (applying) return;
					if (
						// A host save only moves the clean point; it changes nothing to publish.
						(name === 'document-change' &&
							!['remote', 'save'].includes((detail as { kind: string }).kind)) ||
						name === 'document-load'
					) {
						seen = controller.state.document;
						listener();
					}
				}),
				controller.subscribe(check),
			];
			return () => stops.forEach((stop) => stop());
		},
		onIdle: (listener) => controller.subscribe(() => !busy() && listener()),
		busy,
		version: () => controller.documentGeneration,
	};
}

/**
 * Property-level sync of one drawing over a shared session (`ooxml-core/visio/collab`): pages,
 * shapes (per cell, row and section), connectors, masters and media merge per item; styles, theme
 * and document settings are shared whole. Undo and redo act on this
 * window's own edits. It borrows the session; the caller owns its lifetime.
 */
export class DrawingSync {
	readonly #binding: DrawingBinding;
	readonly #stops: (() => void)[] = [];

	constructor(
		drawing: DrawingCollab,
		session: Pick<CollabSession<object>, 'doc' | 'synced' | 'canWrite' | 'on'>,
		private readonly controller: ViewerController,
		options: DrawingSyncOptions,
	) {
		this.#binding = drawing.bindDrawingSession(session, controllerEditor(controller), {
			seed: !options.join,
			onError: (error) => options.fail(error instanceof Error ? error.message : String(error)),
		});
		controller.setSharedHistory(this.#binding);
		this.#stops.push(this.#binding.onHistoryChange(() => controller.refreshHistory()));
	}

	/** Resolves when queued publishing and applying has finished. */
	idle(): Promise<void> {
		return this.#binding.idle();
	}

	dispose(): void {
		for (const stop of this.#stops.splice(0)) stop();
		this.#binding.dispose();
		try {
			this.controller.setSharedHistory(null);
		} catch {
			/* The controller was destroyed first. */
		}
	}
}
