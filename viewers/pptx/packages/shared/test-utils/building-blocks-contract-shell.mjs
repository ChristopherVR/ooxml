/**
 * The shell part of the building-blocks parity contract (`building-blocks-contract.mjs` runs
 * it): the shell state every binding reports (`shellState`, `collaborationState`), joining and
 * leaving a host-owned room through the options, a viewer role taking editing away, and live
 * option updates that must keep the open document (no remount, no reload).
 */

/**
 * A host-owned room on the binding's own `Y.Doc`, with a dependency-free awareness: what a host
 * that owns its transport passes as `collaboration.externalSession`. Already connected and
 * synced, so the blocks' first-write gate opens at once.
 */
export function contractRoom(Y) {
	const doc = new Y.Doc();
	let local = null;
	const events = new Map();
	const emit = (event) => {
		for (const callback of [...(events.get(event) ?? [])]) {
			callback({ added: [], updated: [doc.clientID], removed: [] }, 'local');
		}
	};
	const awareness = {
		clientID: doc.clientID,
		getLocalState: () => local,
		setLocalState: (next) => {
			local = next;
			emit('change');
		},
		setLocalStateField: (key, value) => {
			local = { ...local, [key]: value };
			emit('change');
		},
		getStates: () => new Map(local ? [[doc.clientID, local]] : []),
		on: (event, callback) => {
			if (!events.has(event)) {
				events.set(event, new Set());
			}
			events.get(event).add(callback);
		},
		off: (event, callback) => events.get(event)?.delete(callback),
		destroy: () => undefined,
	};
	const externalSession = {
		doc,
		awareness,
		getSnapshot: () => ({ status: 'connected', synced: true }),
		subscribe: () => () => undefined,
	};
	const config = (role) => ({
		roomId: 'building-blocks-contract',
		serverUrl: '',
		userName: 'Contract host',
		sessionIntent: 'create',
		...(role ? { role } : {}),
		externalSession,
	});
	return { config, destroy: () => doc.destroy() };
}

/**
 * Run the shell steps. `options(extra)` builds the editable options (the contract's callbacks
 * plus `extra`); the deck has unsaved edits (the title "Contract title", one custom show), which
 * every step must keep.
 */
export async function runShellContract({ harness, expect, until, options, Y, onDirtyChange }) {
	const keepsDocument = (label) => {
		const actions = harness.actions();
		expect(harness.state().slideCount, `${label}: slides`).toBe(2);
		expect(harness.state().error, `${label}: error`).toBeNull();
		expect(actions.documentProperties.current.core.title, `${label}: title`).toBe('Contract title');
		expect(actions.customShows.shows, `${label}: custom shows`).toHaveLength(1);
		// A remount or a reload would report a clean deck again.
		expect(onDirtyChange, `${label}: dirty`).not.toHaveBeenLastCalledWith(false);
	};

	// 1. Not sharing: the shell is the host permission, disconnected, nobody connected.
	expect(harness.state().shellState).toEqual({
		canEdit: true,
		status: 'disconnected',
		remoteUsers: [],
		connectedCount: 0,
	});
	expect(harness.state().collaborationState).toMatchObject({
		active: false,
		status: 'off',
		people: [],
	});

	// 2. Live options apply to the mounted canvas and keep the document.
	await harness.update(options({ googleFonts: false, allowExternalContent: false }));
	await harness.settle();
	keepsDocument('live options');
	expect(harness.state().mode).toBe('edit');

	// 3. Joining a host-owned room through the options: connected, synced, still editable.
	const room = contractRoom(Y);
	try {
		await harness.update(options({ collaboration: room.config() }));
		await until(
			() =>
				expect(harness.state().collaborationState).toMatchObject({
					active: true,
					roomId: 'building-blocks-contract',
					status: 'connected',
					synced: true,
				}),
			'the joined room',
		);
		await until(
			() =>
				expect(harness.state().shellState).toMatchObject({
					canEdit: true,
					status: 'connected',
					connectedCount: 1,
				}),
			'the connected shell',
		);
		expect(harness.state().mode).toBe('edit');
		keepsDocument('joined');

		// 4. A viewer role takes editing away (the host permission stays a ceiling).
		await harness.update(options({ collaboration: room.config('viewer') }));
		await until(() => expect(harness.state().shellState.canEdit).toBe(false), 'the viewer role');
		expect(harness.state().mode).toBe('preview');
		expect(harness.actions().availability.insertImage.available).toBe(false);

		// 5. Leaving: off again, editable again, the document kept.
		await harness.update(options({}));
		await until(
			() => expect(harness.state().collaborationState.active).toBe(false),
			'the left room',
		);
		await until(
			() =>
				expect(harness.state().shellState).toEqual({
					canEdit: true,
					status: 'disconnected',
					remoteUsers: [],
					connectedCount: 0,
				}),
			'the disconnected shell',
		);
		expect(harness.state().collaborationState.status).toBe('off');
		expect(harness.state().mode).toBe('edit');
		keepsDocument('left');
	} finally {
		room.destroy();
	}
}
