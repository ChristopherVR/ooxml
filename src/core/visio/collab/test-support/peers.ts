// In-memory peers for the Visio collaboration tests: a package-backed editor on a real collab
// session over the memory transport.
import JSZip from 'jszip';
import { createMemoryHub } from '../../../collab/memory-transport';
import { type CollabSession, createCollabSession } from '../../../collab/session';
import { transportProvider } from '../../../collab/transport-provider';
import { type VisioEdit, editVsdx, parseVsdx } from '../../index';
import {
	type DrawingBinding,
	type DrawingBindingOptions,
	type DrawingEditor,
	bindDrawingSession,
} from '../bind';

export const wait = (ms = 30) => new Promise((resolve) => setTimeout(resolve, ms));
export async function until(check: () => boolean | Promise<boolean>, message = 'condition') {
	for (let i = 0; i < 300 && !(await check()); i++) await wait(10);
	if (!(await check())) throw new Error(`Timed out waiting for ${message}`);
}

/** An editor over package bytes: edits go through the real VSDX writer. */
export function makeEditor(initial: Uint8Array | undefined) {
	let bytes = initial;
	let version = 0;
	let working = false;
	let applied = 0;
	const local = new Set<() => void>();
	const idle = new Set<() => void>();
	const editor: DrawingEditor = {
		exportBytes: () => bytes,
		applyRemote: async (next) => {
			bytes = next;
			version++;
			applied++;
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

/** Waits until both peers have finished publishing and applying. */
export async function settle(...peers: Peer[]) {
	await wait(60);
	for (let round = 0; round < 3; round++) for (const peer of peers) await peer.binding.idle();
	await wait(60);
}
