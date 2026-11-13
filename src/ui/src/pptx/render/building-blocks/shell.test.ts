import { describe, expect, it, vi } from 'vitest';

import { COLLABORATION_OFF } from '../collaboration-state';
import {
	applyBuildingBlockLiveOptions,
	collaborationOptionChange,
	googleFontsOptionChanged,
} from './live-options';
import type { BuildingBlockLiveOptionPort } from './live-options';
import { BUILDING_BLOCK_SHELL_IDLE, createBuildingBlockShell } from './shell';
import type { BuildingBlockShellInput } from './shell';

const user = (clientId: number, cursorX = 0) => ({
	clientId,
	userName: `User ${clientId}`,
	userColor: '#123456',
	activeSlideIndex: 0,
	cursorX,
	cursorY: 0,
	lastUpdated: '',
});

const input = (
	overrides: Partial<BuildingBlockShellInput['shell']> = {},
): BuildingBlockShellInput => ({
	shell: {
		authorizedCanEdit: true,
		configured: true,
		readOnly: false,
		sourcePending: false,
		sourceError: false,
		status: 'connected',
		remoteUsers: [user(2)],
		...overrides,
	},
	collaboration: {
		active: true,
		roomId: 'room',
		status: 'connected',
		synced: true,
		clientId: 1,
		userName: 'Me',
		userColor: '#abcdef',
		remoteUsers: [user(2)],
	},
});

describe('building-block shell', () => {
	it('is idle before a canvas mounts', () => {
		const shell = createBuildingBlockShell(() => null);
		expect(shell.current()).toBe(BUILDING_BLOCK_SHELL_IDLE);
		expect(shell.current().shellState).toEqual({
			canEdit: false,
			status: 'disconnected',
			remoteUsers: [],
			connectedCount: 0,
		});
		expect(shell.current().collaborationState).toBe(COLLABORATION_OFF);
	});

	it('derives the shell projection and the shared state, keeping identity until either changes', () => {
		let current = input();
		const shell = createBuildingBlockShell(() => current);
		const first = shell.current();
		expect(first.shellState).toMatchObject({
			canEdit: true,
			status: 'connected',
			connectedCount: 2,
		});
		expect(first.collaborationState.people.map((person) => person.self)).toEqual([true, false]);
		// A cursor move alone shows nothing new.
		current = { ...input(), shell: { ...input().shell, remoteUsers: [user(2, 40)] } };
		expect(shell.current()).toBe(first);
		current = input({ readOnly: true });
		expect(shell.current()).not.toBe(first);
		expect(shell.current().shellState.canEdit).toBe(false);
	});

	it('passes a state the binding already built through', () => {
		const shell = createBuildingBlockShell(() => ({
			shell: input().shell,
			collaboration: COLLABORATION_OFF,
		}));
		expect(shell.current().collaborationState).toBe(COLLABORATION_OFF);
	});
});

describe('building-block live options', () => {
	const port = (): BuildingBlockLiveOptionPort & Record<string, ReturnType<typeof vi.fn>> => ({
		startCollaboration: vi.fn(),
		stopCollaboration: vi.fn(),
		setGoogleFonts: vi.fn(),
		setAllowExternalContent: vi.fn(),
	});
	const room = { roomId: 'a', serverUrl: '', userName: 'A' };

	it('joins, keeps and leaves a room by the resolved config identity', () => {
		expect(collaborationOptionChange(undefined, room)).toEqual({ kind: 'start', input: room });
		expect(collaborationOptionChange(room, room)).toBeNull();
		expect(collaborationOptionChange(room, { ...room })).toEqual({
			kind: 'start',
			input: { ...room },
		});
		expect(collaborationOptionChange(room, undefined)).toEqual({ kind: 'stop' });
		expect(collaborationOptionChange(undefined, null)).toBeNull();
	});

	it('compares googleFonts by the lock and the mirror it implies', () => {
		expect(googleFontsOptionChanged(false, false)).toBe(false);
		expect(
			googleFontsOptionChanged(
				{ baseUrl: 'https://a.test/css2' },
				{ baseUrl: 'https://a.test/css2' },
			),
		).toBe(false);
		expect(
			googleFontsOptionChanged(
				{ baseUrl: 'https://a.test/css2' },
				{ baseUrl: 'https://b.test/css2' },
			),
		).toBe(true);
		expect(googleFontsOptionChanged(undefined, false)).toBe(true);
	});

	it('pushes only what changed', () => {
		const engine = port();
		applyBuildingBlockLiveOptions({}, {}, engine);
		expect(Object.values(engine).every((fn) => fn.mock.calls.length === 0)).toBe(true);
		applyBuildingBlockLiveOptions(
			{},
			{ googleFonts: false, allowExternalContent: false, collaboration: room },
			engine,
		);
		expect(engine.setGoogleFonts).toHaveBeenCalledWith(false);
		expect(engine.setAllowExternalContent).toHaveBeenCalledWith(false);
		expect(engine.startCollaboration).toHaveBeenCalledWith(room);
		applyBuildingBlockLiveOptions({ collaboration: room }, {}, engine);
		expect(engine.stopCollaboration).toHaveBeenCalledTimes(1);
	});
});
