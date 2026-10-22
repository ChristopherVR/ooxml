// The declarative `collaboration` property of `<docx-editor>`: opens (or borrows) a room session,
// waits for it to be usable, then binds the loaded document through the Yjs mode that
// `startYjsCollaboration` exposes. Also reports the state of any Yjs binding, however started.
import {
	collaborationStateOf,
	onCollaborationStateChange,
	openCollaborationSession,
	sanitizeUserName,
	type CollabSession,
} from 'ooxml-core/collab';
import { docToModel } from './model-adapter';
import { schema } from './schema';
import type { WordYjsOptions } from 'ooxml-core/docx/ui';
import type { EditorCore } from './editor-core';
import { renderDocument } from './editor-render';
import { emit } from './events';
import {
	OFF_STATE,
	type DocxCollaborationOptions,
	type DocxCollaborationState,
} from './collaboration-types';

/** What the controller needs from the editor. */
export interface RoomHost {
	/** The editor is mounted with a view the document can be bound to. */
	ready(): boolean;
	/** A document is being loaded; joining waits for it. */
	loading(): boolean;
	author(): string;
	join(session: CollabSession, options: WordYjsOptions): void;
	leave(): void;
	error(error: Error): void;
	changed(state: DocxCollaborationState): void;
}

interface Opened {
	session: CollabSession;
	owned: boolean;
}

export class DocxRoom {
	#wanted: DocxCollaborationOptions | undefined;
	/** Bumped by every leave, so a join still waiting for its session is abandoned. */
	#epoch = 0;
	#pending: Opened | undefined;
	#mine: Opened | undefined;
	#cancelWait: (() => void) | undefined;
	#watching: { session: CollabSession; stop: () => void } | undefined;

	constructor(private readonly host: RoomHost) {}

	/** The options of the declarative `collaboration` property, if set. */
	get wanted(): DocxCollaborationOptions | undefined {
		return this.#wanted;
	}

	/** Joined or joining through the property. */
	get busy(): boolean {
		return this.#pending !== undefined || this.#mine !== undefined;
	}

	/** The session of a join still waiting to bind, for `reconnectCollaboration`. */
	get pendingSession(): CollabSession | undefined {
		return this.#pending?.session;
	}

	/** Declarative sharing: while set, the open document is shared; undefined leaves. */
	want(options: DocxCollaborationOptions | undefined): void {
		if (options === this.#wanted) return;
		this.#wanted = options;
		this.leave();
		this.resume();
	}

	/** Forgets the property and leaves. True when there was anything to leave. */
	stop(): boolean {
		const had = this.busy || this.#wanted !== undefined;
		this.#wanted = undefined;
		this.leave();
		return had;
	}

	/** Leaves before the document is replaced; true when it was shared through the property. */
	suspend(): boolean {
		const was = this.busy;
		this.leave();
		return was;
	}

	/** Joins the wanted room with the document now open. */
	resume(): void {
		const wanted = this.#wanted;
		if (!wanted || this.busy || !this.host.ready() || this.host.loading()) return;
		this.#join(wanted).catch((cause: unknown) => this.host.error(asError(cause)));
	}

	/** The Yjs binding started or stopped (through the property or `startYjsCollaboration`). */
	yjsChanged(session: CollabSession | undefined): void {
		if (session) this.#watch(session, false);
		else {
			const mine = this.#mine;
			this.#mine = undefined;
			this.#unwatch();
			if (mine?.owned) mine.session.destroy();
		}
		this.#changed();
	}

	state(): DocxCollaborationState {
		return this.#watching ? collaborationStateOf(this.#watching.session) : OFF_STATE;
	}

	async #join(options: DocxCollaborationOptions): Promise<void> {
		const epoch = this.#epoch;
		const opened = openCollaborationSession(options, {
			defaultName: sanitizeUserName(this.host.author(), 'Author'),
			broadcastPrefix: 'ooxml-docx:',
		});
		this.#pending = opened;
		this.#watch(opened.session, opened.owned);
		this.#changed();
		try {
			const create = options.intent !== 'join';
			await this.#until(opened.session, create);
			if (epoch !== this.#epoch) return;
			if (!this.host.ready()) {
				this.leave();
				return;
			}
			this.#pending = undefined;
			this.#mine = opened;
			this.host.join(opened.session, {
				documentId: options.roomId || opened.session.roomId,
				initializeIfEmpty: create,
				...(create && !opened.session.synced ? { seedWithoutSync: true } : {}),
			});
		} catch (cause) {
			if (epoch === this.#epoch) this.leave();
			throw cause;
		}
	}

	/** Resolves once the room exchanged state, or (for a creator) once the write gate opened. */
	#until(session: CollabSession, create: boolean): Promise<void> {
		return new Promise<void>((resolve) => {
			const stops: Array<() => void> = [];
			const done = () => {
				for (const stop of stops) stop();
				this.#cancelWait = undefined;
				resolve();
			};
			const check = () => {
				if (session.synced || (create && session.gate.isOpen())) done();
			};
			this.#cancelWait = done;
			stops.push(session.on('synced', check), session.on('ready', check));
			check();
		});
	}

	/** Leaves the room; a session the host passed in stays alive, one made here is destroyed. */
	leave(): void {
		this.#epoch++;
		this.#cancelWait?.();
		const pending = this.#pending;
		this.#pending = undefined;
		if (this.#mine) {
			try {
				this.host.leave();
			} finally {
				const mine = this.#mine as Opened | undefined;
				this.#mine = undefined;
				if (mine?.owned) mine.session.destroy();
			}
		}
		if (pending) {
			this.#unwatch();
			if (pending.owned) pending.session.destroy();
			this.#changed();
		}
	}

	#watch(session: CollabSession, reportErrors: boolean): void {
		if (this.#watching?.session === session) return;
		this.#unwatch();
		const stops = [onCollaborationStateChange(session, () => this.#changed())];
		if (reportErrors) stops.push(session.on('error', (error) => this.host.error(error)));
		this.#watching = {
			session,
			stop: () => {
				for (const stop of stops) stop();
			},
		};
	}

	#unwatch(): void {
		this.#watching?.stop();
		this.#watching = undefined;
	}

	#changed(): void {
		this.host.changed(this.state());
	}
}

/** Binds the open document to a synchronized session (the body of `startYjsCollaboration`). */
export function joinYjs(core: EditorCore, session: CollabSession, options: WordYjsOptions): void {
	if (!core.view) throw new Error('Mount and load the document before starting collaboration.');
	core.collab.startYjs(session, core.view.state.doc, {
		...options,
		initialMedia: new Map([...(options.initialMedia ?? []), ...core.inserts.pendingMedia]),
		initialComments: options.initialComments ?? core.model.comments ?? [],
	});
	core.loadGeneration++;
	core.detachedState = undefined;
	renderDocument(core);
	core.roleChanged();
}

/** Detaches the Yjs binding, keeping its current content, comments and new picture parts. */
export function leaveYjs(core: EditorCore): void {
	const yjs = core.collab.yjs;
	if (yjs) {
		core.model = docToModel(yjs.state(schema).doc, core.model);
		if (yjs.sharedComments) core.model = { ...core.model, comments: yjs.comments.all() };
		for (const [name, part] of yjs.media.all()) core.inserts.pendingMedia.set(name, part);
	}
	core.collab.stop();
	core.detachedState = undefined;
	if (core.element.isConnected) renderDocument(core);
	core.roleChanged();
}

/** The controller of one editor, reporting through its events and error channel. */
export function roomFor(core: EditorCore): DocxRoom {
	return new DocxRoom({
		ready: () => Boolean(core.view),
		loading: () => core.loading,
		author: () => core.reviewAuthor,
		join: (session, options) => joinYjs(core, session, options),
		leave: () => leaveYjs(core),
		error: (error) => emit(core.element, 'document-error', error),
		changed: (state) => emit(core.element, 'collaboration-change', state),
	});
}

function asError(error: unknown): Error {
	return error instanceof Error ? error : new Error(String(error));
}
