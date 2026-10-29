import { toBase64 } from 'ooxml-core/collab';
import { createChatStore } from 'ooxml-core/teams';
import * as Y from 'yjs';
import { defineTeamsApp, type TeamsApp } from './teams-app';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

/** A stored workspace snapshot; `version` is the schema stamp it carries (none: a current one). */
function seedSnapshot(room: string, version?: number): void {
	const doc = new Y.Doc();
	createChatStore(doc, { id: 'ada', name: 'Ada' }).createChannel({
		id: 'general',
		name: 'General',
	});
	if (version !== undefined) doc.getMap('teams:meta').set('schemaVersion', version);
	localStorage.setItem(`teams:doc:${room}`, toBase64(Y.encodeStateAsUpdate(doc)));
}

async function mount(room: string): Promise<TeamsApp> {
	const el = document.createElement('teams-app') as TeamsApp;
	el.setAttribute('workspace-id', room);
	el.setAttribute('user-name', 'Ada');
	el.setAttribute('user-id', 'ada');
	el.setAttribute('server-config', LOCAL);
	document.body.append(el);
	await el.updateComplete;
	await tick(150);
	await el.updateComplete;
	return el;
}

const composer = (el: TeamsApp): HTMLElement & { disabled: boolean } =>
	el.shadowRoot!.querySelector('office-ui-chat-composer') as HTMLElement & { disabled: boolean };

describe('<teams-app> with a workspace written by a newer version', () => {
	it('shows a banner and disables composing', async () => {
		seedSnapshot('lock-newer', 999);
		const el = await mount('lock-newer');
		expect(el.client!.getState().schema.state).toBe('unsupported-schema');
		const banner = el.shadowRoot!.querySelector('[data-schema-banner]');
		expect(banner?.getAttribute('role')).toBe('alert');
		expect(banner?.textContent).toMatch(/newer version of OpenTeams/);
		expect(composer(el).disabled).toBe(true);
	});

	it('does not open the create channel dialog', async () => {
		seedSnapshot('lock-dialogs', 999);
		const el = await mount('lock-dialogs');
		(el as unknown as { createChannel: () => void }).createChannel();
		await el.updateComplete;
		expect(el.creatingChannel).toBe(false);
	});

	it('stays editable, with no banner, for a current workspace', async () => {
		seedSnapshot('lock-current');
		const el = await mount('lock-current');
		expect(el.client!.getState().schema.state).toBe('ok');
		expect(el.shadowRoot!.querySelector('[data-schema-banner]')).toBeNull();
		expect(composer(el).disabled).toBe(false);
	});
});
