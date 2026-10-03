import type { OfficeProfile } from 'ooxml-ui/controls';
import type { PresenceParticipant } from 'ooxml-ui/presence';
import type { ViewerController, ViewerState } from './controller.js';

type Collab = typeof import('ooxml-core/collab');
type Session = ReturnType<Collab['createCollabSession']>;
type Presence = HTMLElement & { participants: PresenceParticipant[] };

/** Room names the core accepts: 1-64 letters, digits, hyphens or underscores. */
const ROOM = /^[A-Za-z0-9_-]{1,64}$/;

/**
 * File > Share: a live session over `ooxml-core/collab` (Yjs). Peers in other windows of this
 * browser share the whole VSDX package through a BroadcastChannel; remote changes arrive as one
 * undoable history step. The collab runtime loads only when sharing starts.
 */
export class ViewerShare {
	readonly #panel: HTMLElement;
	readonly #room: HTMLInputElement;
	readonly #start: HTMLButtonElement;
	readonly #stop: HTMLButtonElement;
	readonly #status: HTMLElement;
	readonly #people: Presence;
	#session: Session | null = null;
	#teardown: (() => void) | null = null;
	#revision = 0;
	/** The package this window last published or adopted, to skip repeats of it. */
	#last: Uint8Array | null = null;
	#adopting = false;
	#pending: Uint8Array | null = null;
	#starting = false;
	constructor(
		root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly profile: () => OfficeProfile,
		private readonly loadCollab: () => Promise<Collab> = () => import('ooxml-core/collab'),
	) {
		this.#panel = root.querySelector('.share-panel')!;
		this.#room = this.#panel.querySelector('.share-room')!;
		this.#start = this.#panel.querySelector('[data-share="start"]')!;
		this.#stop = this.#panel.querySelector('[data-share="stop"]')!;
		this.#status = this.#panel.querySelector('.share-status')!;
		this.#people = this.#panel.querySelector<Presence>('office-ui-presence')!;
		this.#room.value = `visio-${Math.random().toString(36).slice(2, 8)}`;
	}
	get active(): boolean {
		return this.#session !== null;
	}
	wire(): () => void {
		const Abort = this.#panel.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		this.#panel.addEventListener(
			'click',
			(event) => {
				const action = (event.target as Element).closest?.<HTMLElement>('[data-share]');
				if (action?.dataset.share === 'start') void this.start(this.#room.value.trim());
				if (action?.dataset.share === 'stop') this.stop();
			},
			{ signal: events.signal },
		);
		return () => {
			events.abort();
			this.stop();
		};
	}
	async start(roomId: string): Promise<void> {
		if (this.#session || this.#starting) return;
		if (!ROOM.test(roomId))
			return this.#say('Use 1 to 64 letters, digits, hyphens or underscores for the session name.');
		this.#starting = true;
		this.#render();
		try {
			const collab = await this.loadCollab();
			if (!collab.canBroadcast())
				return this.#say('Sharing needs BroadcastChannel, which this browser does not provide.');
			this.#join(collab, roomId);
		} catch (error) {
			this.#say(error instanceof Error ? error.message : String(error));
		} finally {
			this.#starting = false;
			this.#render();
		}
	}
	stop(): void {
		this.#teardown?.();
		this.#teardown = null;
		this.#session?.destroy();
		this.#session = null;
		this.#pending = null;
		this.#revision = 0;
		this.#last = null;
		this.#people.participants = [];
		this.#say('Not sharing.');
		this.#render();
	}
	#join(collab: Collab, roomId: string): void {
		const adapter = collab.packageAdapter();
		const me = this.profile();
		const session = collab.createCollabSession({
			roomId: `visio-${roomId}`,
			provider: collab.transportProvider({
				transport: collab.createBroadcastTransport({ roomId: `visio-${roomId}` }),
			}),
			user: { name: me.displayName || 'Visio user', color: me.avatarColor },
			teardown: true,
		});
		this.#session = session;
		const publish = () => {
			const bytes = this.#localBytes();
			if (!bytes || !session.canWrite()) return;
			adapter.write(session.doc, { bytes, revision: this.#revision }, collab.LOCAL_ORIGIN);
			this.#revision = adapter.read(session.doc).revision;
			this.#last = bytes;
		};
		const adopt = () => {
			if (adapter.isEmpty(session.doc)) return;
			const model = adapter.read(session.doc);
			// Two windows can write the same revision at once; the bytes tell them apart.
			if (model.revision < this.#revision) return;
			if (model.revision === this.#revision && same(model.bytes, this.#last)) return;
			this.#revision = model.revision;
			this.#last = model.bytes;
			this.#apply(model.bytes);
		};
		const map = session.doc.getMap('package');
		const observer = (_event: unknown, transaction: { origin: unknown }) => {
			if (transaction.origin !== collab.LOCAL_ORIGIN) adopt();
		};
		map.observe(observer);
		// The room wins when it already has a drawing; otherwise this window seeds it.
		const settle = () => (adapter.isEmpty(session.doc) ? publish() : adopt());
		const stopReady = session.on('ready', settle);
		const stopPeers = session.on('peers', (peers) => this.#showPeers(peers));
		const stopEvents = this.controller.onEvent((name, detail) => {
			if (name === 'document-change' && (detail as { kind: string }).kind !== 'remote') publish();
			if (name === 'document-load' && !this.#adopting) publish();
		});
		const stopState = this.controller.subscribe((state) => this.#flush(state));
		this.#teardown = () => {
			map.unobserve(observer);
			stopReady();
			stopPeers();
			stopEvents();
			stopState();
		};
		if (session.canWrite()) settle();
		this.#showPeers(session.peers());
		this.#say(`Sharing as ${me.displayName || 'Visio user'} in session ${roomId}.`);
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
			if (error) this.#say(error instanceof Error ? error.message : String(error));
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
	}
	#showPeers(peers: readonly { clientId: number; userName: string; userColor: string }[]): void {
		const me = this.profile();
		this.#people.participants = [
			{ id: 'self', name: me.displayName || 'Visio user', color: me.avatarColor, self: true },
			...peers.map((peer) => ({
				id: String(peer.clientId),
				name: peer.userName,
				color: peer.userColor,
			})),
		];
	}
	#say(message: string): void {
		this.#status.textContent = message;
	}
	#render(): void {
		const active = this.active;
		this.#start.hidden = active;
		this.#stop.hidden = !active;
		this.#start.disabled = this.#starting;
		this.#room.disabled = active || this.#starting;
	}
}

function same(a: Uint8Array, b: Uint8Array | null): boolean {
	if (!b || a.byteLength !== b.byteLength) return false;
	for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
	return true;
}
