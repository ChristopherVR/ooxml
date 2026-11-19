// In-memory peers for the Visio collaboration tests: a package-backed editor on a real collab
// session over the memory transport.
import JSZip from 'jszip';
import * as Y from 'yjs';
import { createMemoryHub } from '../../../collab/memory-transport';
import { type CollabSession, createCollabSession } from '../../../collab/session';
import { transportProvider } from '../../../collab/transport-provider';
import { type VisioEdit, editVsdx, parseVsdx } from '../../index';
import {
	type DrawingBinding,
	type DrawingChange,
	type DrawingBindingOptions,
	type DrawingEditor,
	bindDrawingSession,
} from '../bind';

export const wait = (ms = 30) => new Promise((resolve) => setTimeout(resolve, ms));
/** How long the peers may take, by the clock: a loaded runner stretches every step alike. */
const PATIENCE_MS = 20_000;
export async function until(check: () => boolean | Promise<boolean>, message = 'condition') {
	const deadline = Date.now() + PATIENCE_MS;
	while (!(await check())) {
		if (Date.now() > deadline) throw new Error(`Timed out waiting for ${message}`);
		await wait(10);
	}
}

/** An editor over package bytes: edits go through the real VSDX writer. */
export function makeEditor(initial: Uint8Array | undefined) {
	let bytes = initial;
	let version = 0;
	let working = false;
	let applied = 0;
	/** The `change` handed to each `applyRemote` (undefined: the whole package). */
	const changes: (DrawingChange | undefined)[] = [];
	const local = new Set<() => void>();
	const idle = new Set<() => void>();
	const editor: DrawingEditor = {
		exportBytes: () => bytes,
		applyRemote: async (next, change) => {
			bytes = next;
			version++;
			applied++;
			changes.push(change);
		},
		onLocalChange: (listener) => {
			local.add(listener);
			return () => local.delete(listener);
		},
		onIdle: (listener) => {
			idle.add(listener);
			return () => idle.delete(listener);
		},
		busy: () => working,
		version: () => version,
	};
	const announce = () => {
		version++;
		for (const listener of [...local]) listener();
		for (const listener of [...idle]) listener();
	};
	return {
		editor,
		get applied() {
			return applied;
		},
		changes,
		get bytes() {
			return bytes!;
		},
		async edit(...edits: VisioEdit[]) {
			working = true;
			try {
				bytes = (await editVsdx(bytes!, edits)).bytes;
			} finally {
				working = false;
			}
			announce();
		},
		/** Replaces the package by `change(zip)` as a local edit (for parts the edit API lacks). */
		async patch(change: (zip: JSZip) => Promise<void> | void) {
			working = true;
			try {
				const zip = await JSZip.loadAsync(bytes!);
				await change(zip);
				bytes = await zip.generateAsync({ type: 'uint8array' });
			} finally {
				working = false;
			}
			announce();
		},
		model: async () => parseVsdx(bytes!),
		async text(path: string): Promise<string | undefined> {
			return (await JSZip.loadAsync(bytes!)).file(path)?.async('string');
		},
		async binary(path: string): Promise<Uint8Array | undefined> {
			return (await JSZip.loadAsync(bytes!)).file(path)?.async('uint8array');
		},
	};
}

export function createRoom() {
	const hub = createMemoryHub();
	const sessions: CollabSession<object>[] = [];
	const bindings: DrawingBinding[] = [];
	const join = (name: string, initial?: Uint8Array, options: DrawingBindingOptions = {}) => {
		const session = createCollabSession({
			roomId: 'drawing',
			provider: transportProvider({ transport: hub.createTransport('drawing') }),
			user: { name },
			heartbeatMs: 0,
			syncGraceMs: 0,
		});
		sessions.push(session);
		const side = makeEditor(initial);
		const errors: unknown[] = [];
		const binding = bindDrawingSession(session, side.editor, {
			onError: (e) => errors.push(e),
			...options,
		});
		bindings.push(binding);
		return { session, binding, errors, side };
	};
	return {
		hub,
		join,
		destroy() {
			for (const binding of bindings.splice(0)) binding.dispose();
			for (const session of sessions.splice(0)) session.destroy();
		},
	};
}
export type Peer = ReturnType<ReturnType<typeof createRoom>['join']>;

/** The updates a peer's room holds. */
const room = (peer: Peer): string => Y.encodeStateVector(peer.session.doc).join();
/** What a peer has done so far: its room and the drawings it applied or edited. */
const progress = (peer: Peer): string => `${room(peer)}|${peer.side.editor.version()}`;

/**
 * Waits until the peers have nothing left to do: a whole round of `idle()` (and of the timers
 * it schedules) changes no room or drawing, and every room holds the same updates. Delivery over
 * the memory hub is synchronous, so this follows the bindings' own signals instead of sleeping.
 */
export async function settle(...peers: Peer[]) {
	const deadline = Date.now() + PATIENCE_MS;
	for (;;) {
		const before = peers.map(progress);
		await Promise.all(peers.map((peer) => peer.binding.idle()));
		await wait(0);
		const still = peers.every((peer, i) => progress(peer) === before[i]);
		if (still && new Set(peers.map(room)).size === 1) return;
		if (Date.now() > deadline) throw new Error('Timed out waiting for the peers to settle');
	}
}

/** Waits until the peer's drawing is in the room, so a peer that joins next adopts it. */
export async function published(peer: Peer) {
	await until(() => peer.session.doc.getMap('visio:meta').size > 0, 'publish');
	await settle(peer);
}
