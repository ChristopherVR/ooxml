import { describe, expect, it, vi } from 'vitest';

import {
	buildCollaborationState,
	COLLABORATION_OFF,
	collaborationStatesEqual,
	createCollaborationStateReporter,
} from './collaboration-state';

const remote = { clientId: 2, userName: 'Ada', userColor: '#ef4444' };

describe('buildCollaborationState', () => {
	it('is the frozen off state when inactive', () => {
		const state = buildCollaborationState({
			active: false,
			roomId: 'r',
			status: 'connected',
			synced: true,
			remoteUsers: [remote],
		});
		expect(state).toBe(COLLABORATION_OFF);
		expect(Object.isFrozen(state)).toBe(true);
		expect(state).toMatchObject({ active: false, status: 'off', synced: false, people: [] });
	});

	it('lists the local person first, then the remote ones', () => {
		const state = buildCollaborationState({
			active: true,
			roomId: 'room-1',
			status: 'connected',
			synced: true,
			clientId: 1,
			userName: 'Me',
			userColor: '#3b82f6',
			remoteUsers: [remote],
		});
		expect(state).toEqual({
			active: true,
			roomId: 'room-1',
			status: 'connected',
			synced: true,
			people: [
				{ clientId: 1, name: 'Me', color: '#3b82f6', self: true },
				{ clientId: 2, name: 'Ada', color: '#ef4444', self: false },
			],
		});
	});

	it('omits the local person until its client id is known and derives a colour', () => {
		const pending = buildCollaborationState({
			active: true,
			status: 'connecting',
			synced: false,
			clientId: null,
			remoteUsers: [],
		});
		expect(pending.people).toEqual([]);
		expect(pending.roomId).toBeUndefined();
		const known = buildCollaborationState({
			active: true,
			status: 'connecting',
			synced: false,
			clientId: 7,
			userName: 'Me',
			remoteUsers: [],
		});
		expect(known.people[0]?.color).toMatch(/^#[0-9a-f]{6}$/u);
	});
});

describe('collaborationStatesEqual and the reporter', () => {
	const base = {
		active: true,
		roomId: 'r',
		status: 'connected' as const,
		synced: true,
		clientId: 1,
		userName: 'Me',
		userColor: '#3b82f6',
		remoteUsers: [remote],
	};

	it('compares structurally', () => {
		expect(
			collaborationStatesEqual(buildCollaborationState(base), buildCollaborationState(base)),
		).toBe(true);
		expect(
			collaborationStatesEqual(
				buildCollaborationState(base),
				buildCollaborationState({ ...base, synced: false }),
			),
		).toBe(false);
		expect(
			collaborationStatesEqual(
				buildCollaborationState(base),
				buildCollaborationState({ ...base, remoteUsers: [] }),
			),
		).toBe(false);
	});

	it('reports only real changes and stays silent while off', () => {
		const report = vi.fn();
		const reporter = createCollaborationStateReporter(report);
		reporter.update(COLLABORATION_OFF);
		expect(report).not.toHaveBeenCalled();
		reporter.update(buildCollaborationState(base));
		reporter.update(buildCollaborationState(base));
		expect(report).toHaveBeenCalledTimes(1);
		reporter.update(COLLABORATION_OFF);
		expect(report).toHaveBeenCalledTimes(2);
		expect(reporter.current()).toBe(COLLABORATION_OFF);
	});
});
