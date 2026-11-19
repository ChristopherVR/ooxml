// Real-time co-editing of `<xlsx-editor>`: joins a `ooxml-core/collab` session, binds the edit
// session through `bindWorkbookSession`, publishes the selection as presence and reports peers and
// connection changes. While bound, undo and redo go through the binding's Yjs undo manager (only
// this window's edits); the edit session's own undo would publish a snapshot restore as an edit.
// The Yjs runtime loads only when sharing starts.
import type { CollabSession, ConnectionStatus } from 'ooxml-core/collab';
import type { CellRange, EditSession, RemoteRange } from 'ooxml-core/xlsx';
import type { WorkbookBinding, XlsxPresence } from 'ooxml-core/xlsx/collab';
import type { EditHistory, Selection, SelectionModel } from 'ooxml-core/xlsx/ui';
import { createRoomSession, type CollabRuntime } from './collaboration-session';
import type { EditorCore } from './editor-core';
import { emit } from './events';
import {
	OFF_STATE,
	type XlsxCollaborationOptions,
	type XlsxCollaborationState,
} from './collaboration-types';

const loadRuntime = (): Promise<CollabRuntime> =>
	Promise.all([import('ooxml-core/collab'), import('ooxml-core/xlsx/collab')]);

/** What the controller needs from the editor. */
export interface CollaborationHost {
	edit(): EditSession | undefined;
	readonly selection: SelectionModel;
	userName(): string;
	t(key: string, vars?: Record<string, string | number>): string;
	toast(message: string, kind: 'info' | 'warning' | 'error'): void;
	/** Reports a failure as `workbook-error`. */
	error(error: Error): void;
	/** Announces `collaboration-change` and refreshes the title bar. */
	changed(state: XlsxCollaborationState): void;
	/** Locks the editor while the session role is `viewer`; independent of the host's `readOnly`. */
	setRoleReadOnly(locked: boolean): void;
}

interface Active {
	collab: CollabRuntime[0];
	session: CollabSession<XlsxPresence>;
	binding: WorkbookBinding;
	owned: boolean;
	stop: Array<() => void>;
}

/** The part of the selection a collaborator sees: the range holding the active cell. */
export function presenceRange(selection: Selection): CellRange {
	const { row, col } = selection.active;
	const holding = selection.ranges.find(
		(r) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col,
	);
	return (
		holding ??
		selection.ranges[selection.ranges.length - 1] ?? {
			start: selection.active,
			end: selection.active,
		}
	);
}

export class XlsxCollaboration {
	#active: Active | undefined;
	#starting: Promise<void> | undefined;
	/** Bumped by every stop, so a join still loading the runtime is abandoned. */
	#epoch = 0;
	#wanted: XlsxCollaborationOptions | undefined;
	/** The options of the room joined or being joined. */
	#joined: XlsxCollaborationOptions | undefined;
	/** The room left by `park`, joined again by `resume`. */
	#parked: XlsxCollaborationOptions | undefined;
	#lost = false;
	#connected = false;
	readonly #listeners = new Set<() => void>();

	constructor(
		private readonly host: CollaborationHost,
		private readonly load: () => Promise<CollabRuntime> = loadRuntime,
	) {}

	get active(): boolean {
		return this.#active !== undefined;
	}

	/** Joined or joining. */
	get busy(): boolean {
		return this.#active !== undefined || this.#starting !== undefined;
	}

	/** The options of the declarative `collaboration` property, if set. */
	get wanted(): XlsxCollaborationOptions | undefined {
		return this.#wanted;
	}

	/**
	 * Declarative sharing: while set, the open workbook is shared in that room, and a workbook
	 * opened later joins it too (the room's content wins when it already has one). Undefined stops.
	 */
	want(options: XlsxCollaborationOptions | undefined): void {
		if (options === this.#wanted) return;
		this.#wanted = options;
		this.#leave();
		this.resume();
	}

	/** Leaves the room before the workbook is replaced; true when it was shared. */
	suspend(): boolean {
		const was = this.busy;
		this.#leave();
		return was;
	}

	/**
	 * Leaves the room while the editor is off the page (the session's timers would otherwise keep it
	 * alive), remembering it for `resume`. A session the host passed in stays alive.
	 */
	park(): void {
		const joined = this.#joined;
		this.#leave();
		this.#parked = joined;
	}

	/** Joins the wanted room, or the one `park` left, with the workbook now open. */
	resume(): void {
		const wanted = this.#wanted ?? this.#parked;
		if (!wanted || this.busy || !this.host.edit()) return;
		this.#parked = undefined;
		this.start(wanted).catch((error: unknown) => {
			const failure = asError(error);
			this.host.toast(this.host.t(failure.message), 'error');
			this.host.error(failure);
		});
	}

	/** Joins a room. Rejects when one is active, no workbook is open or the room is invalid. */
	start(options: XlsxCollaborationOptions): Promise<void> {
		if (this.#active || this.#starting)
			return Promise.reject(new Error('Stop the current collaboration session first.'));
		const starting = this.#start(options).finally(() => {
			if (this.#starting !== starting) return;
			this.#starting = undefined;
			if (!this.#active) this.#joined = undefined;
		});
		this.#starting = starting;
		this.#joined = options;
		return starting;
	}

	async #start(options: XlsxCollaborationOptions): Promise<void> {
		if (!this.host.edit()) throw new Error('Open or create a workbook before sharing it.');
		const epoch = this.#epoch;
		const [collab, xlsx] = await this.load();
		if (epoch !== this.#epoch) return;
		const edit = this.host.edit();
		if (!edit) throw new Error('Open or create a workbook before sharing it.');
		const session =
			options.session ?? createRoomSession(collab, xlsx, options, this.host.userName());
		let leftOutShown = false;
		const binding = xlsx.bindWorkbookSession(session, edit, {
			onError: (error) => this.host.error(asError(error)),
			// Once per session: a dismissed notice does not return with every remote change.
			onLeftOut: () => {
				if (leftOutShown) return;
				leftOutShown = true;
				this.host.toast(this.host.t(collab.LEFT_OUT_MESSAGE), 'warning');
			},
		});
		const active: Active = { collab, session, binding, owned: !options.session, stop: [] };
		this.#active = active;
		this.#lost = false;
		this.#connected = session.status === 'connected';
		const publish = (selection: Selection) =>
			binding.setSelection(selection.sheet, presenceRange(selection));
		active.stop.push(
			this.host.selection.onChange(publish),
			session.on('peers', () => this.#changed()),
			session.on('synced', () => this.#changed()),
			session.on('status', (status) => this.#status(status)),
			session.on('error', (error) => {
				// A refused update stops sharing; say why in the user's language.
				if (collab.isUpdateRefusal(error)) this.host.toast(this.host.t(error.message), 'error');
				this.host.error(error);
			}),
			edit.onChange((change) => {
				// Sheets a peer added, moved or removed shift where remote selections land.
				if (change.external) this.#notify();
			}),
		);
		this.host.setRoleReadOnly(session.role === 'viewer');
		publish(this.host.selection.get());
		this.#changed();
	}

	/**
	 * Leaves the room and forgets the `collaboration` property. A session the host passed in stays
	 * alive; one the editor made is destroyed.
	 */
	stop(): void {
		this.#wanted = undefined;
		this.#leave();
	}

	#leave(): void {
		this.#epoch++;
		this.#starting = undefined;
		this.#joined = this.#parked = undefined;
		const active = this.#active;
		if (!active) return;
		this.#active = undefined;
		for (const stop of active.stop) stop();
		active.binding.dispose();
		if (active.owned) active.session.destroy();
		this.host.setRoleReadOnly(false);
		this.#changed();
	}

	/** Restarts the connection, keeping the shared workbook and queued changes. */
	reconnect(): void {
		this.#active?.session.reconnect();
	}

	/** The binding's history while shared (only this window's edits), else undefined. */
	history(): EditHistory | undefined {
		const binding = this.#active?.binding;
		if (!binding) return undefined;
		return {
			undo: () => binding.undo(),
			redo: () => binding.redo(),
			canUndo: () => binding.canUndo(),
			canRedo: () => binding.canRedo(),
			undoLabel: () => undefined,
			redoLabel: () => undefined,
		};
	}

	remoteSelections(): readonly RemoteRange[] {
		return this.#active?.binding.remoteSelections() ?? [];
	}

	onRemoteChange(listener: () => void): () => void {
		this.#listeners.add(listener);
		return () => {
			this.#listeners.delete(listener);
		};
	}

	state(): XlsxCollaborationState {
		const active = this.#active;
		return active ? active.collab.collaborationStateOf(active.session) : OFF_STATE;
	}

	#status(status: ConnectionStatus): void {
		const { t } = this.host;
		if (status === 'connected') {
			if (this.#lost)
				this.host.toast(t('Reconnected. Changes made while offline were sent.'), 'info');
			this.#lost = false;
			this.#connected = true;
		} else if (
			(status === 'disconnected' || status === 'error') &&
			this.#connected &&
			!this.#lost &&
			// A refusal was already explained, and the connection does not come back by itself.
			!this.#active?.session.refusal
		) {
			this.#lost = true;
			this.host.toast(
				t('Connection lost. Keep working; your changes are sent when it returns.'),
				'warning',
			);
		}
		this.#changed();
	}

	#notify(): void {
		for (const listener of [...this.#listeners]) listener();
	}

	#changed(): void {
		this.#notify();
		this.host.changed(this.state());
	}
}

/** The controller of one editor, reporting through its toasts, events and render loop. */
export function collaborationFor(core: EditorCore): XlsxCollaboration {
	const { host } = core;
	return new XlsxCollaboration({
		edit: () => core.session,
		selection: core.selection,
		userName: () => core.authorName,
		t: core.ctx.t,
		toast: (message, kind) => core.shell?.toast(message, kind),
		error: (error) => emit(host, 'workbook-error', { error, message: core.ctx.t(error.message) }),
		changed: (state) => {
			emit(host, 'collaboration-change', state);
			core.requestRender();
		},
		setRoleReadOnly: (locked) => core.setRoleReadOnly(locked),
	});
}

function asError(error: unknown): Error {
	return error instanceof Error ? error : new Error(String(error));
}
