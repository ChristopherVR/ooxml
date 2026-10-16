import { defineTeamsApp, type TeamsApp } from './teams-app';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

async function mount(name: string, room: string): Promise<TeamsApp> {
	const el = document.createElement('teams-app') as TeamsApp;
	el.setAttribute('workspace-id', room);
	el.setAttribute('user-name', name);
	el.setAttribute('user-id', name.toLowerCase());
	el.setAttribute('server-config', LOCAL);
	document.body.append(el);
	await el.updateComplete;
	return el;
}

/** Ada and Bob in one workspace, Bob having just mentioned Ada. */
async function mentioned(room: string) {
	const ada = await mount('Ada', room);
	const bob = await mount('Bob', room);
	await tick(200);
	await ada.client!.send({ text: 'hello' });
	await tick(200);
	await bob.client!.send({
		text: 'ping @Ada',
		mentions: [{ kind: 'user', id: 'ada', name: 'Ada' }],
	});
	await tick(250);
	await ada.updateComplete;
	return { ada, bob };
}

describe('Activity in <teams-app>', () => {
	it('puts Activity first in the app rail with an unread badge, and lists the mention', async () => {
		const { ada } = await mentioned('app-activity-feed');
		const rail = ada.shadowRoot!.querySelector('office-ui-app-rail') as HTMLElement & {
			items: { id: string; badge?: number }[];
		};
		expect(rail.items[0]).toMatchObject({ id: 'activity', badge: 1 });
		rail.dispatchEvent(new CustomEvent('office-rail-select', { detail: { id: 'activity' } }));
		await ada.updateComplete;
		const feed = ada.shadowRoot!.querySelector('office-ui-activity-feed') as HTMLElement;
		await (feed as unknown as { updateComplete: Promise<boolean> }).updateComplete;
		expect(feed.shadowRoot!.querySelector('.line')?.textContent).toBe(
			'Bob mentioned you in # General',
		);
	});

	it('opens an item on its message and marks it read', async () => {
		const { ada } = await mentioned('app-activity-open');
		const id = ada.client!.getState().activity[0]!.id;
		ada.rail = 'activity';
		await ada.updateComplete;
		const feed = ada.shadowRoot!.querySelector('office-ui-activity-feed')!;
		feed.dispatchEvent(new CustomEvent('office-activity-open', { detail: { id } }));
		await tick(100);
		await ada.updateComplete;
		expect(ada.client!.getState().activityUnread).toBe(0);
		expect(ada.highlightId).not.toBe('');
		expect(ada.mobileView).toBe('conversation');
	});

	it('toasts a new mention, but not during Do not disturb', async () => {
		const room = 'app-activity-toast';
		const ada = await mount('Ada', room);
		const bob = await mount('Bob', room);
		await tick(200);
		await bob.client!.send({ text: 'hi' });
		await tick(150);
		const mention = { kind: 'user', id: 'ada', name: 'Ada' } as const;
		ada.client!.setAvailability('dnd');
		await tick(50);
		await bob.client!.send({ text: 'quiet @Ada', mentions: [mention] });
		await tick(250);
		await ada.updateComplete;
		expect(ada.shadowRoot!.querySelector('.activity-toast')).toBeNull();
		ada.client!.setAvailability('available');
		await tick(50);
		await bob.client!.send({ text: 'loud @Ada', mentions: [mention] });
		await tick(250);
		await ada.updateComplete;
		const toast = ada.shadowRoot!.querySelector('.activity-toast');
		expect(toast?.textContent).toContain('Bob mentioned you in # General');
		expect(toast?.textContent).toContain('loud Ada');
	});

	it('changes the notification level from Settings', async () => {
		const ada = await mount('Ada', 'app-activity-settings');
		ada.settingsOpen = true;
		await ada.updateComplete;
		const settings = ada.shadowRoot!.querySelector('teams-settings') as HTMLElement & {
			category: string;
			updateComplete: Promise<boolean>;
		};
		settings.category = 'notifications';
		await settings.updateComplete;
		const mentionsOnly = settings.shadowRoot!.querySelector<HTMLInputElement>(
			'input[name="notification-level"][value="mentions"]',
		)!;
		mentionsOnly.checked = true;
		mentionsOnly.dispatchEvent(new Event('change'));
		await tick(50);
		expect(ada.client!.getState().notificationSettings.level).toBe('mentions');
		expect(settings.shadowRoot!.querySelector('#panel-notifications')!.textContent).not.toContain(
			'not available yet',
		);
	});
});
