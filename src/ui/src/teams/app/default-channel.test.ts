import { DEFAULT_CHANNEL, type TeamsState } from 'ooxml-core/teams';
import { describe, expect, it } from 'vitest';
import { teamsI18n } from '../i18n/translator';
import { localizeDefaultChannel } from './controller';

const state = (name: string, topic: string): TeamsState =>
	({
		channels: [
			{ ...DEFAULT_CHANNEL, name, topic },
			{ id: 'c2', name: 'General', topic: '' },
		],
		channel: { ...DEFAULT_CHANNEL, name, topic },
	}) as unknown as TeamsState;

describe('the starting channel in the display language', () => {
	it('shows the seeded name and topic translated, here and as the selected channel', () => {
		const shown = localizeDefaultChannel(
			state(DEFAULT_CHANNEL.name, DEFAULT_CHANNEL.topic),
			teamsI18n('fr'),
		);
		expect(shown.channels[0]).toMatchObject({
			name: 'Général',
			topic: "Annonces pour toute l'entreprise",
		});
		expect(shown.channel).toMatchObject({ name: 'Général' });
		// Another channel someone named "General" is their content, never translated.
		expect(shown.channels[1]!.name).toBe('General');
	});

	it('keeps a name or topic someone changed, and English as stored', () => {
		const renamed = state('Announcements', DEFAULT_CHANNEL.topic);
		expect(localizeDefaultChannel(renamed, teamsI18n('de')).channels[0]).toMatchObject({
			name: 'Announcements',
			topic: 'Unternehmensweite Ankündigungen',
		});
		const english = state(DEFAULT_CHANNEL.name, DEFAULT_CHANNEL.topic);
		expect(localizeDefaultChannel(english, teamsI18n('en'))).toBe(english);
	});
});
