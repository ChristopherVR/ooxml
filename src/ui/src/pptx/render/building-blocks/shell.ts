/**
 * The shell state every binding's building blocks report beside `actions`: the collaboration
 * shell projection (`shellState`, the shape a custom status bar reads) and the shared
 * `CollaborationState` (`collaborationState`, what `collaborationChange` carries). Each binding
 * reads the raw inputs from its own editor; the derivation, the idle value and the identity rule
 * (a new object only when something a shell shows changed) are here, once.
 *
 * @module render/building-blocks/shell
 */
import type {
	CollaborationShellInput,
	CollaborationShellState,
} from '../collaboration-shell-state';
import { resolveCollaborationShellState } from '../collaboration-shell-state';
import type { CollaborationState, CollaborationStateInput } from '../collaboration-state';
import {
	buildCollaborationState,
	COLLABORATION_OFF,
	collaborationStatesEqual,
} from '../collaboration-state';

export interface BuildingBlockShell {
	/** Effective edit permission plus connection and presence counts (`describeCollaborationShellState`). */
	readonly shellState: CollaborationShellState;
	/** The shared `CollaborationState` of the session the blocks show; `COLLABORATION_OFF` when none. */
	readonly collaborationState: CollaborationState;
}

export interface BuildingBlockShellInput {
	readonly shell: CollaborationShellInput;
	/** The state the binding already reports (`collaborationChange`), or the input to build it from. */
	readonly collaboration: CollaborationState | CollaborationStateInput;
}

/** What the blocks report before a canvas mounts: view-only, not sharing. */
export const BUILDING_BLOCK_SHELL_IDLE: BuildingBlockShell = Object.freeze({
	shellState: Object.freeze({
		canEdit: false,
		status: 'disconnected',
		remoteUsers: Object.freeze([]) as unknown as CollaborationShellState['remoteUsers'],
		connectedCount: 0,
	}) as CollaborationShellState,
	collaborationState: COLLABORATION_OFF,
});

const isState = (
	value: CollaborationState | CollaborationStateInput,
): value is CollaborationState => 'people' in value;

export function describeBuildingBlockShell(input: BuildingBlockShellInput): BuildingBlockShell {
	return {
		shellState: resolveCollaborationShellState(input.shell),
		collaborationState: isState(input.collaboration)
			? input.collaboration
			: buildCollaborationState(input.collaboration),
	};
}

function shellStatesEqual(a: CollaborationShellState, b: CollaborationShellState): boolean {
	return (
		a.canEdit === b.canEdit &&
		a.status === b.status &&
		a.connectedCount === b.connectedCount &&
		a.remoteUsers.length === b.remoteUsers.length &&
		a.remoteUsers.every((user, index) => {
			const other = b.remoteUsers[index];
			return (
				other !== undefined &&
				user.clientId === other.clientId &&
				user.userName === other.userName &&
				user.userColor === other.userColor &&
				user.role === other.role &&
				user.activeSlideIndex === other.activeSlideIndex &&
				user.selectedElementId === other.selectedElementId
			);
		})
	);
}

/** True when two shells show the same thing. */
export function buildingBlockShellsEqual(a: BuildingBlockShell, b: BuildingBlockShell): boolean {
	return (
		a === b ||
		(shellStatesEqual(a.shellState, b.shellState) &&
			collaborationStatesEqual(a.collaborationState, b.collaborationState))
	);
}

/**
 * A view of the shell over `read`, returning the previous object while nothing a shell shows
 * changed (cursor moves alone keep it), so a `computed`, a signal or a getter re-renders only on
 * real changes. `read` returns null before a canvas mounts.
 */
export function createBuildingBlockShell(read: () => BuildingBlockShellInput | null): {
	current(): BuildingBlockShell;
} {
	let last: BuildingBlockShell = BUILDING_BLOCK_SHELL_IDLE;
	return {
		current() {
			const input = read();
			const next = input ? describeBuildingBlockShell(input) : BUILDING_BLOCK_SHELL_IDLE;
			if (!buildingBlockShellsEqual(last, next)) {
				last = next;
			}
			return last;
		},
	};
}
