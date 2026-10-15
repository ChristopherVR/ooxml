// @vitest-environment node
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createTeams } from 'ooxml-ui/teams';

// Every binding exposes the raw client from `createTeams`, so the Chat and Calendar actions and
// state reach React `useTeams`, Vue `useTeams`, Solid `createTeamsClient`, Svelte `teamsStore` and
// Angular `TeamsService` without each binding listing them.
const rawHooks = {
	react: 'react/src/index.ts',
	vue: 'vue/src/index.ts',
	solid: 'solid/src/index.ts',
	svelte: 'svelte/src/store.ts',
	angular: 'angular/src/service.ts',
} as const;

describe.each(Object.entries(rawHooks))('%s raw hook', (_name, file) => {
	it('hands out the core client built by createTeams', () => {
		const source = readFileSync(resolve(process.cwd(), 'packages', file), 'utf8');
		expect(source).toContain('createTeams(');
	});
});

describe('the raw client', () => {
	it('carries the Chat and Calendar API', async () => {
		const client = createTeams({
			workspaceId: `raw-hook-${Date.now()}`,
			user: { id: 'ada', name: 'Ada' },
			reminderCheckMs: 0,
		});
		try {
			for (const action of [
				'openDirectChat',
				'createGroupChat',
				'renameChat',
				'addToChat',
				'removeFromChat',
				'leaveChat',
				'createMeeting',
				'updateMeeting',
				'cancelMeeting',
				'meetingIcs',
				'joinMeeting',
			] as const)
				expect(typeof client[action], action).toBe('function');
			const state = client.getState();
			expect(state.chats).toEqual([]);
			expect(state.meetings).toEqual([]);
			expect(Array.isArray(state.contacts)).toBe(true);
			expect(typeof state.viewerZone).toBe('string');
		} finally {
			client.destroy();
		}
	});
});
