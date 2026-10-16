import { html } from 'lit';
import { teamsI18n, type TeamsI18n, type TeamsKey } from '../i18n/translator';

export const SETTINGS_CATEGORIES = [
	['general', 'settings.general'],
	['appearance', 'settings.appearance'],
	['notifications', 'settings.notifications'],
	['files', 'settings.files'],
	['connection', 'settings.connection'],
] as const satisfies readonly (readonly [string, TeamsKey])[];

export type SettingsCategory = (typeof SETTINGS_CATEGORIES)[number][0];
export interface SettingTopic {
	category: SettingsCategory;
	title: string;
	description: string;
}
// Index only controls that are implemented in the settings surface.
const TOPICS: readonly {
	category: SettingsCategory;
	title: TeamsKey;
	description: TeamsKey;
}[] = [
	{ category: 'general', title: 'settings.profile', description: 'settings.topic.profile' },
	{ category: 'appearance', title: 'settings.theme', description: 'settings.topic.theme' },
	{ category: 'appearance', title: 'settings.density', description: 'settings.topic.density' },
	{ category: 'appearance', title: 'settings.appNames', description: 'settings.appNamesText' },
	{
		category: 'notifications',
		title: 'settings.activityNotifications',
		description: 'settings.topic.activityNotifications',
	},
	{ category: 'notifications', title: 'settings.sound', description: 'settings.soundText' },
	{
		category: 'notifications',
		title: 'settings.desktop',
		description: 'settings.topic.desktop',
	},
	{
		category: 'notifications',
		title: 'settings.followed',
		description: 'settings.topic.followed',
	},
	{ category: 'files', title: 'settings.fileOpen', description: 'settings.topic.fileOpen' },
	{
		category: 'connection',
		title: 'settings.connection',
		description: 'settings.topic.connection',
	},
];

/** The label of a settings category, in the display language. */
export const categoryLabel = (category: SettingsCategory, i18n: TeamsI18n): string =>
	i18n.t(SETTINGS_CATEGORIES.find(([id]) => id === category)![1]);

export function settingsSearch(
	query: string,
	select: (topic: SettingTopic) => void,
	i18n: TeamsI18n = teamsI18n(),
) {
	const { t } = i18n;
	const words = query.trim().toLocaleLowerCase(i18n.locale).split(/\s+/u);
	const topics: SettingTopic[] = TOPICS.map((topic) => ({
		category: topic.category,
		title: t(topic.title),
		description: t(topic.description),
	}));
	const matches = topics.filter((topic) => {
		const text =
			`${categoryLabel(topic.category, i18n)} ${topic.title} ${topic.description}`.toLocaleLowerCase(
				i18n.locale,
			);
		return words.every((word) => text.includes(word));
	});
	return html`<section class="settings-results" aria-label=${t('settings.resultsLabel')}>
		<h2>${t('settings.results')}</h2>
		<p role="status">
			${matches.length ? i18n.tn('settings.found', matches.length) : t('settings.noMatch')}
		</p>
		<ul>
			${matches.map(
				(topic) =>
					html`<li>
						<button type="button" @click=${() => select(topic)}>
							<strong>${topic.title}</strong><span>${categoryLabel(topic.category, i18n)}</span
							><small>${topic.description}</small>
						</button>
					</li>`,
			)}
		</ul>
	</section>`;
}
