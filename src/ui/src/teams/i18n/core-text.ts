// The core (`ooxml-core/teams`) words its notices and errors in English. The UI shows them in the
// display locale by matching the known messages and swapping in the dictionary's wording. A
// message this table does not know (a server's own error text, say) is shown as written.
import type { MessageParams } from 'ooxml-core/i18n';
import type { TeamsI18n, TeamsKey } from './translator';

type Rule = readonly [
	pattern: RegExp,
	key: TeamsKey,
	params?: (match: RegExpExecArray) => MessageParams,
];

const exact = (text: string): RegExp =>
	new RegExp(`^${text.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')}$`, 'u');

const RULES: readonly Rule[] = [
	[exact('That person is not available to chat with'), 'core.personUnavailable'],
	[exact('Pick at least one person'), 'core.pickPerson'],
	[exact('The chat could not be created'), 'core.chatNotCreated'],
	[
		exact('The meeting needs a title and an end after its start (24 hours at most)'),
		'core.meetingInvalid',
	],
	[exact('Only the organizer can change a meeting'), 'core.organizerChange'],
	[exact('Only the organizer can cancel a meeting'), 'core.organizerCancel'],
	[
		exact('Only the organizer can cancel an occurrence of a meeting'),
		'core.organizerCancelOccurrence',
	],
	[
		exact('Only the organizer can change an occurrence of a meeting'),
		'core.organizerChangeOccurrence',
	],
	[exact('Camera unavailable'), 'core.cameraUnavailable'],
	[
		exact('Reattach or discard the missing draft attachments before sending'),
		'core.reattachDrafts',
	],
	[exact('The edit was not applied; the draft was retained'), 'core.editNotApplied'],
	[exact('Message send canceled. The message was saved in Drafts.'), 'core.sendCanceled'],
	[
		/^Message was not sent: (.*)\. Saved in Drafts\.$/su,
		'core.sendFailed',
		(m) => ({ reason: m[1] ?? '' }),
	],
	[exact('This message cannot be pinned'), 'core.cannotPin'],
	[exact('Could not open the file'), 'core.openFailed'],
	[exact('This meeting is not available'), 'core.meetingUnavailable'],
	[exact('The channel is no longer available'), 'core.channelGone'],
	[
		exact('The channel is no longer available; the uploaded files were not shared'),
		'core.channelGoneUploaded',
	],
	[exact('Configure file storage before sharing files'), 'core.configureStorage'],
	[exact('Configure file storage before sharing attachments'), 'core.configureStorage'],
	[exact('The files could not be shared'), 'core.filesNotShared'],
	[/^Choose between 1 and (\d+) files$/u, 'core.chooseFiles', (m) => ({ max: m[1] ?? '' })],
	[exact('Enter a workbook name'), 'core.workbookName'],
	[/^Upload failed \((\d+)\)$/u, 'core.uploadFailed', (m) => ({ status: m[1] ?? '' })],
	[exact('Server returned an invalid file URL'), 'core.invalidFileUrl'],
	[exact('This file cannot be saved to the channel'), 'core.cannotSaveFile'],
	[exact('The file could not be uploaded'), 'core.fileNotUploaded'],
	[/^Could not load content \((\d+)\)$/u, 'core.loadFailed', (m) => ({ status: m[1] ?? '' })],
	[exact('Content streaming is unavailable'), 'core.streamingUnavailable'],
	[exact('Content is too large to preview'), 'core.tooLarge'],
	[exact('The file is no longer available'), 'core.fileGone'],
	[/^Call is full \((\d+) participants\)$/u, 'core.callFull', (m) => ({ max: m[1] ?? '' })],
	[/^Link failed \((\d+)\)$/u, 'core.linkFailed', (m) => ({ status: m[1] ?? '' })],
];

/** A core notice or error message in the display locale; unknown text is returned unchanged. */
export function localizeCoreText(text: string, i18n: TeamsI18n): string {
	if (i18n.locale === 'en') return text;
	for (const [pattern, key, params] of RULES) {
		const match = pattern.exec(text);
		if (match) return i18n.t(key, params?.(match));
	}
	return text;
}
