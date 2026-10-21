import type { VisioDocument } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller';

export type Collab = typeof import('ooxml-core/collab');
export type Session = ReturnType<Collab['createCollabSession']>;

export interface PackageSyncOptions {
	/** `join` never seeds an empty room with the drawing this window already has open. */
	join: boolean;
	/** Reports an adoption failure. */
	fail(message: string): void;
}

/**
 * Whole-package sync of one drawing over a shared session: the VSDX bytes live in
 * `session.doc.getMap('package')` and the last writer wins. Remote changes arrive as one undoable
 * history step. It borrows the session; the caller owns its lifetime.
 */
export class PackageSync {
	readonly #adapter: ReturnType<Collab['packageAdapter']>;
	readonly #stops: (() => void)[] = [];
	#revision = 0;
	/** The package this window last published or adopted, to skip repeats of it. */
	#last: Uint8Array | null = null;
	/** The drawing this window last published or adopted, so a local load or edit is sent once. */
	#shared: VisioDocument | null = null;
	#adopting = false;
	#pending: Uint8Array | null = null;
	/** Publishes a local drawing the room has not seen; set once the session has settled. */
	#catchUp: (() => void) | null = null;

	constructor(
		private readonly collab: Collab,
		private readonly session: Session,
		private readonly controller: ViewerController,
		private readonly options: PackageSyncOptions,
	) {
		this.#adapter = collab.packageAdapter();
		const map = session.doc.getMap('package');
		const observer = (_event: unknown, transaction: { origin: unknown }) => {
			if (transaction.origin !== collab.LOCAL_ORIGIN) this.#adopt();
		};
		map.observe(observer);
		// The room wins when it already has a drawing; otherwise this window seeds it. From then on a
		// settled local drawing the room has not seen is published from state, not only from the
		// load and change events: the controller drops an event for later listeners when an earlier
		// one changes state (a host fitting the page on load), and sharing must not depend on order.
		const settle = () => {
			this.#catchUp = () => this.#publish();
			if (!this.#adapter.isEmpty(session.doc)) this.#adopt();
			else if (options.join) this.#shared = this.controller.state.document;
			else this.#publish();
		};
		this.#stops.push(
			() => map.unobserve(observer),
			session.on('ready', settle),
			this.controller.onEvent((name, detail) => {
				const kind = name === 'document-change' ? (detail as { kind: string }).kind : null;
				// A host save only moves the clean point; it changes nothing to publish.
				if (kind !== null && kind !== 'remote' && kind !== 'save') this.#publish();
				if (name === 'document-load' && !this.#adopting) this.#publish();
			}),
			this.controller.subscribe((state) => this.#flush(state)),
		);
		if (session.canWrite()) settle();
	}

	dispose(): void {
		for (const stop of this.#stops.splice(0)) stop();
		this.#catchUp = null;
		this.#pending = null;
	}

	#publish(): void {
		const { document } = this.controller.state;
		if (this.#adopting || this.#pending || document === this.#shared || !this.session.canWrite())
			return;
		const bytes = this.#localBytes();
		if (!bytes) return;
		const { doc } = this.session;
		this.#adapter.write(doc, { bytes, revision: this.#revision }, this.collab.LOCAL_ORIGIN);
		this.#revision = this.#adapter.read(doc).revision;
		this.#last = bytes;
		this.#shared = document;
	}

	#adopt(): void {
		if (this.#adapter.isEmpty(this.session.doc)) return;
		const model = this.#adapter.read(this.session.doc);
		// Two windows can write the same revision at once; the bytes tell them apart.
		if (model.revision < this.#revision) return;
		if (model.revision === this.#revision && same(model.bytes, this.#last)) return;
		this.#revision = model.revision;
		this.#last = model.bytes;
		this.#apply(model.bytes);
	}

	#localBytes(): Uint8Array | null {
		const { edit, loading } = this.controller.state;
		if (!edit.sourceAvailable || edit.busy || loading) return null;
		try {
			return this.controller.exportVsdx().bytes;
		} catch {
			return null;
		}
	}

	#apply(bytes: Uint8Array): void {
		const { edit, loading, document } = this.controller.state;
		if (edit.busy || loading) {
			this.#pending = bytes;
			return;
		}
		this.#pending = null;
		const done = (error?: unknown) => {
			this.#adopting = false;
			if (error) this.options.fail(error instanceof Error ? error.message : String(error));
			else this.#shared = this.controller.state.document;
		};
		this.#adopting = true;
		const run =
			document && edit.sourceAvailable
				? this.controller.applyRemoteSource(bytes)
				: this.controller.load(bytes);
		run.then(() => done(), done);
	}

	#flush(state: ViewerState): void {
		if (this.#pending && !state.edit.busy && !state.loading) this.#apply(this.#pending);
		else this.#catchUp?.();
	}
}

function same(a: Uint8Array, b: Uint8Array | null): boolean {
	if (!b || a.byteLength !== b.byteLength) return false;
	for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
	return true;
}
