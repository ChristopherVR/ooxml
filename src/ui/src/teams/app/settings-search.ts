import { html } from 'lit';

export const SETTINGS_CATEGORIES = [
	['general', 'General'],
	['appearance', 'Appearance and accessibility'],
	['notifications', 'Notifications and activity'],
	['files', 'Files and links'],
	['connection', 'Connection'],
] as const;

export type SettingsCategory = (typeof SETTINGS_CATEGORIES)[number][0];
interface SettingTopic {
	category: SettingsCategory;
	title: string;
	description: string;
}
// Index only controls that are implemented in the settings surface.
const TOPICS: SettingTopic[] = [
	{ category: 'general', title: 'Profile', description: 'Your name and workspace on this device' },
	{
		category: 'appearance',
		title: 'Theme',
		description: 'Follow system, light or dark appearance',
	},
	{
		category: 'appearance',
		title: 'Chat density',
		description: 'Comfy or compact message spacing',
	},
	{
		category: 'appearance',
		title: 'Show app names',
		description: 'Display labels beneath app bar icons',
	},
	{
		category: 'notifications',
		title: 'Followed threads',
		description: 'Threads I start and threads I reply to',
	},
	{
		category: 'files',
		title: 'File open preference',
		description: 'Open Word, PowerPoint and Excel in OpenTeams or browser',
	},
	{
		category: 'connection',
		title: 'Connection',
		description: 'Workspace server, sync URL, signaling URL, access token and ICE servers',
	},
];

export function settingsSearch(query: string, select: (topic: SettingTopic) => void) {
	const words = query.trim().toLocaleLowerCase().split(/\s+/u);
	const matches = TOPICS.filter((topic) => {
		const category = SETTINGS_CATEGORIES.find(([id]) => id === topic.category)![1];
		const text = `${category} ${topic.title} ${topic.description}`.toLocaleLowerCase();
		return words.every((word) => text.includes(word));
	});
	return html`<section class="settings-results" aria-label="Settings search results">
		<h2>Search results</h2>
		<p role="status">
			${matches.length ? `${matches.length} ${matches.length === 1 ? 'setting' : 'settings'} found` : 'No matching settings. Try another keyword.'}
		</p>
		<ul>
			${matches.map(
				(topic) =>
					html`<li>
						<button type="button" @click=${() => select(topic)}>
							<strong>${topic.title}</strong
							><span>${SETTINGS_CATEGORIES.find(([id]) => id === topic.category)![1]}</span
							><small>${topic.description}</small>
						</button>
					</li>`,
			)}
		</ul>
	</section>`;
}
