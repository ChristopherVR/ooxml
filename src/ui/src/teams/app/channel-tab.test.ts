// @vitest-environment jsdom
import { defineTeamsApp, type TeamsApp } from './teams-app';
import type { TeamsChannelTab } from './channel-tab';
import type { TeamsContentPreview } from './content-preview';

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
	vi.restoreAllMocks();
});
const tick = () => new Promise((resolve) => setTimeout(resolve, 10));

async function mount(): Promise<TeamsApp> {
	const app = document.createElement('teams-app') as TeamsApp;
	app.userName = 'Ada';
	app.userId = 'ada';
	app.workspaceId = crypto.randomUUID();
	app.config = { mode: 'local', iceServers: [] };
	document.body.append(app);
	await app.updateComplete;
	app.client!.createChannel('Project');
	await tick();
	await app.updateComplete;
	return app;
}

describe('channel tab workspace', () => {
	it('opens shared website tabs and keeps a file preview stable during unrelated state changes', async () => {
		const app = await mount();
		const tab = app.client!.addTab('Project page', { type: 'website', url: 'https://site.test/' })!;
		app.tab = tab.id;
		await tick();
		await app.updateComplete;
		const pane = app.shadowRoot!.querySelector('teams-channel-tab') as TeamsChannelTab;
		await pane.updateComplete;
		await tick();
		const detail = pane.detail;
		expect(detail?.url).toBe('https://site.test/');
		app.client!.setAvailability('busy');
		await tick();
		await app.updateComplete;
		await pane.updateComplete;
		expect(pane.detail).toBe(detail);
		app.client!.removeTab(tab.id);
		await tick();
		await app.updateComplete;
		expect(app.shadowRoot!.querySelector('teams-channel-tab')).toBe(pane);
		expect(app.shadowRoot!.textContent).toContain('Your open copy remains here');
		expect(pane.detail).toBe(detail);
		(app.shadowRoot!.querySelector('[role="tab"]') as HTMLButtonElement).click();
		await app.updateComplete;
		expect(app.shadowRoot!.querySelector('office-ui-chat-composer')).not.toBeNull();
	});
	it('does not leave a workbook preview when the user keeps unsaved edits', async () => {
		const app = await mount();
		app.previewContent({
			attachment: { name: 'site.html', kind: 'other' },
			url: 'https://site.test/',
		});
		await app.updateComplete;
		const preview = app.shadowRoot!.querySelector('teams-content-preview') as TeamsContentPreview;
		preview.dirty = true;
		const answer = async (button: 'confirm' | 'cancel') => {
			const dialog = await vi.waitFor(() => {
				const found = preview.shadowRoot!.querySelector('office-ui-message-dialog');
				if (!found) throw new Error('no dialog yet');
				return found;
			});
			(dialog.shadowRoot!.querySelector(`button.${button}`) as HTMLButtonElement).click();
			await new Promise((resolve) => setTimeout(resolve, 0));
		};
		app
			.shadowRoot!.querySelector('office-ui-app-rail')!
			.dispatchEvent(new CustomEvent('office-rail-select', { detail: { id: 'files' } }));
		await answer('cancel');
		expect(app.preview).not.toBeNull();
		expect(app.rail).toBe('teams');
		app
			.shadowRoot!.querySelector('office-ui-app-rail')!
			.dispatchEvent(new CustomEvent('office-rail-select', { detail: { id: 'files' } }));
		// Discarding clears the edits and carries on with the navigation that was held back.
		await answer('confirm');
		expect(app.preview).toBeNull();
		expect(app.rail).toBe('files');
	});
});
