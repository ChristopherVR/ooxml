import type { CollaborationOptions, CollaborationState } from 'ooxml-core/collab';
import type { OfficeProfile } from '../controls';
import type { PresenceParticipant } from '../presence';
import type { ViewerController } from './controller';
import { type Collab, PackageSync, type Session } from './viewer-share-sync';

type Presence = HTMLElement & { participants: PresenceParticipant[] };

/** Not sharing: what `collaborationState` reports before a join and after a leave. */
export const NOT_SHARING: CollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});

interface Joined {
	collab: Collab;
	session: Session;
	owned: boolean;
	sync: PackageSync;
	stops: (() => void)[];
	/** The viewer was editable before a viewer role made it read-only. */
	restoreCanEdit: boolean;
}

/**
 * Sharing of one `<visio-viewer>` over `ooxml-core/collab` (Yjs): the File > Share panel and the
 * host's `collaboration` property drive the same session. Peers share the whole VSDX package
 * (see `PackageSync`). The collab runtime loads only when sharing starts.
 */
export class ViewerShare {
	readonly #panel: HTMLElement;
	readonly #room: HTMLInputElement;
	readonly #start: HTMLButtonElement;
	readonly #stop: HTMLButtonElement;
	readonly #status: HTMLElement;
	readonly #people: Presence;
	#joined: Joined | null = null;
	#starting = false;
	/** Bumped by every leave, so a join still loading the runtime is abandoned. */
	#epoch = 0;
	#wanted: CollaborationOptions | null = null;
	/** Called whenever `state` may have changed. */
	onChange: (state: CollaborationState) => void = () => {};
	/** Called with the failure of a declarative join. */
	onError: (error: Error) => void = () => {};
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
		return this.#joined !== null;
	}
	/** The options of the declarative `collaboration` property, if set. */
	get wanted(): CollaborationOptions | null {
		return this.#wanted;
	}
	get state(): CollaborationState {
		const joined = this.#joined;
		return joined ? joined.collab.collaborationStateOf(joined.session) : NOT_SHARING;
	}
	/** Declarative sharing: while set, the drawing is shared; null leaves the room. */
	want(options: CollaborationOptions | null): void {
		if (options === this.#wanted) return;
		this.#wanted = options;
		this.#leave();
		this.#resume();
	}
	/** Leaves the room and forgets the `collaboration` property. */
	stop(): void {
		this.#wanted = null;
		this.#leave();
	}
	reconnect(): void {
		this.#joined?.session.reconnect();
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
		// A reconnected viewer rejoins the room its host asked for.
		this.#resume();
		return () => {
			events.abort();
			// The property outlives a disconnect; the session does not.
			this.#leave();
		};
	}
	/** The panel's Start: a room of this browser's own. */
	async start(roomId: string): Promise<void> {
		if (this.#joined || this.#starting) return;
		try {
			await this.#open({ roomId });
		} catch (error) {
			this.#say(error instanceof Error ? error.message : String(error));
		}
	}
	#resume(): void {
		const wanted = this.#wanted;
		if (!wanted || this.#joined || this.#starting) return;
		this.#open(wanted).catch((error: unknown) => {
			const failure = error instanceof Error ? error : new Error(String(error));
			this.#say(failure.message);
			this.onError(failure);
		});
	}
	async #open(options: CollaborationOptions): Promise<void> {
		const epoch = this.#epoch;
		this.#starting = true;
		this.#render();
		try {
			const collab = await this.loadCollab();
			if (epoch !== this.#epoch) return;
			this.#join(collab, options);
		} finally {
			if (epoch === this.#epoch) this.#starting = false;
			this.#render();
		}
	}
	#join(collab: Collab, options: CollaborationOptions): void {
		const me = this.profile();
		// Panel rooms keep the wire name they always had (`ooxml-collab:visio-<room>`), so windows on
		// an older build and rooms started from the panel stay interoperable.
		const { session, owned } = collab.openCollaborationSession(
			{ ...options, user: { color: me.avatarColor, ...options.user } },
			{ defaultName: me.displayName || 'Visio user', broadcastPrefix: 'ooxml-collab:visio-' },
		);
		const sync = new PackageSync(collab, session, this.controller, {
			join: options.intent === 'join',
			fail: (message) => this.#say(message),
		});
		// A viewer role is read-only: its edits are never published, so they are not offered either.
		const readOnly = session.role === 'viewer' && this.controller.canEdit;
		if (readOnly) this.controller.canEdit = false;
		this.#joined = {
			collab,
			session,
			owned,
			sync,
			restoreCanEdit: readOnly,
			stops: [
				session.on('peers', () => this.#changed()),
				session.on('synced', () => this.#changed()),
				session.on('status', () => this.#changed()),
			],
		};
		this.#say(`Sharing as ${session.identity.userName} in session ${session.roomId}.`);
		this.#changed();
	}
	#leave(): void {
		this.#epoch++;
		this.#starting = false;
		const joined = this.#joined;
		if (!joined) {
			this.#render();
			return;
		}
		this.#joined = null;
		for (const stop of joined.stops) stop();
		joined.sync.dispose();
		// A session the host passed in stays alive; one this viewer made is destroyed.
		if (joined.owned) joined.session.destroy();
		if (joined.restoreCanEdit) this.controller.canEdit = true;
		this.#say('Not sharing.');
		this.#changed();
	}
	#changed(): void {
		this.#showPeers();
		this.#render();
		this.onChange(this.state);
	}
	#showPeers(): void {
		const joined = this.#joined;
		if (!joined) {
			this.#people.participants = [];
			return;
		}
		this.#people.participants = [
			{
				id: 'self',
				name: joined.session.identity.userName,
				color: joined.session.identity.userColor,
				self: true,
			},
			...joined.session.peers().map((peer) => ({
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
