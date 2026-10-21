import type { CollaborationState } from 'ooxml-core/collab';

import { assignUserColor, type ConnectionStatus } from './collaboration-presence';

export type { CollaborationState };

/** What every binding already tracks about the session it shows. */
export interface CollaborationStateInput {
	/** A session is joined (viewer-owned or host-owned). False yields {@link COLLABORATION_OFF}. */
	active: boolean;
	roomId?: string | undefined;
	status: ConnectionStatus;
	/** The first exchange with the room finished (the first-write gate is open). */
	synced: boolean;
	/** The local awareness client id; the local person is omitted until it is known. */
	clientId?: number | null | undefined;
	userName?: string | undefined;
	userColor?: string | undefined;
	/** Remote presences, as each binding sanitises them for cursors and avatars. */
	remoteUsers: readonly { clientId: number; userName: string; userColor: string }[];
}

/** The state of an editor that is not sharing. One frozen object, so `===` tells it apart. */
export const COLLABORATION_OFF: CollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: Object.freeze([]) as unknown as CollaborationState['people'],
});

/** Build the shared `CollaborationState` a host reads from the PowerPoint viewer. */
export function buildCollaborationState(input: CollaborationStateInput): CollaborationState {
	if (!input.active) return COLLABORATION_OFF;
	const name = input.userName ?? '';
	const people: CollaborationState['people'] = [];
	if (input.clientId !== null && input.clientId !== undefined) {
		people.push({
			clientId: input.clientId,
			name,
			color: input.userColor || assignUserColor(name),
			self: true,
		});
	}
	for (const user of input.remoteUsers) {
		people.push({
			clientId: user.clientId,
			name: user.userName,
			color: user.userColor,
			self: false,
		});
	}
	return {
		active: true,
		roomId: input.roomId || undefined,
		status: input.status,
		synced: input.synced,
		people,
	};
}

/** True when two states report the same thing, so a binding only emits real changes. */
export function collaborationStatesEqual(a: CollaborationState, b: CollaborationState): boolean {
	if (a === b) return true;
	if (
		a.active !== b.active ||
		a.roomId !== b.roomId ||
		a.status !== b.status ||
		a.synced !== b.synced ||
		a.people.length !== b.people.length
	) {
		return false;
	}
	return a.people.every((person, index) => {
		const other = b.people[index];
		return (
			other !== undefined &&
			person.clientId === other.clientId &&
			person.name === other.name &&
			person.color === other.color &&
			person.self === other.self
		);
	});
}

/**
 * Emits the state to `report` when it differs from the last one reported. Start with
 * {@link COLLABORATION_OFF} as the baseline, so an editor that never shares stays silent.
 */
export function createCollaborationStateReporter(report: (state: CollaborationState) => void): {
	update: (state: CollaborationState) => void;
	current: () => CollaborationState;
} {
	let last = COLLABORATION_OFF;
	return {
		update: (state) => {
			if (collaborationStatesEqual(last, state)) return;
			last = state;
			report(state);
		},
		current: () => last,
	};
}
