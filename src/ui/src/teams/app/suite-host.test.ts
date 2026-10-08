import { defineTeamsFileActions, type TeamsFileActions } from './file-actions';
import { defineTeamsChannelTab, type TeamsChannelTab } from './channel-tab';
import type { TeamsClient, FileEntry } from 'ooxml-core/teams';

beforeAll(() => {
	defineTeamsFileActions();
	defineTeamsChannelTab();
});
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

it('allows a host to download local attachments without fetching a shell URL', async () => {
	const actions = document.createElement('teams-file-actions') as TeamsFileActions;
	actions.file = { name: 'brief.docx', url: 'https://suite.test/#/file/one' } as FileEntry;
	const fileUrl = vi.fn();
	actions.client = { fileUrl } as unknown as TeamsClient;
	const intercepted = vi.fn((event: Event) => event.preventDefault());
	actions.addEventListener('teams-file-transfer', intercepted);
	document.body.append(actions);
	await actions.updateComplete;
	const button = [...actions.shadowRoot!.querySelectorAll('button')].find(
		(b) => b.textContent?.trim() === 'Download',
	)!;
	button.click();
	await actions.updateComplete;
	expect(intercepted).toHaveBeenCalledOnce();
	expect(fileUrl).not.toHaveBeenCalled();
});

it('routes a pinned document through the same suite file opener', async () => {
	const tab = document.createElement('teams-channel-tab') as TeamsChannelTab;
	tab.tab = {
		id: 'tab',
		channelId: 'general',
		name: 'Brief',
		content: {
			type: 'file',
			attachment: { name: 'brief.docx', kind: 'docx', url: 'https://suite.test/#/file/one' },
		},
		createdAt: 0,
		createdBy: 'user',
	};
	tab.client = { fileUrl: async () => 'https://suite.test/#/file/one' } as unknown as TeamsClient;
	const open = vi.fn((event: Event) => event.preventDefault());
	tab.addEventListener('teams-open-file', open);
	document.body.append(tab);
	await tab.updateComplete;
	await new Promise((resolve) => setTimeout(resolve, 0));
	await tab.updateComplete;
	expect(open).toHaveBeenCalledOnce();
	expect(tab.shadowRoot!.textContent).toContain('Open document');
	expect(tab.shadowRoot!.querySelector('teams-content-preview')).toBeNull();
});
