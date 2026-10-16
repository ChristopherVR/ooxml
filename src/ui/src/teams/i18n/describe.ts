// Text the core builds in English from its own data, rebuilt here in the display locale.
import type { ActivityItem } from 'ooxml-core/teams';
import type { TeamsI18n } from './translator';

/** The line an Activity item shows ("Ada mentioned you in # General"), in the display locale. */
export function describeActivityIn(item: ActivityItem, i18n: TeamsI18n): string {
	const where = item.chat ? item.channelName : `# ${item.channelName}`;
	const actor = item.actorName;
	switch (item.kind) {
		case 'mention':
			return i18n.t('activity.mention', { actor, where });
		case 'reply':
			return i18n.t('activity.reply', { actor, where });
		case 'reaction':
			return item.emoji
				? i18n.t('activity.reactionEmoji', { actor, emoji: item.emoji })
				: i18n.t('activity.reaction', { actor });
		case 'pin':
			return i18n.t('activity.pin', { actor, where });
		case 'meeting':
			return i18n.t('activity.meeting', { title: item.channelName });
	}
}
