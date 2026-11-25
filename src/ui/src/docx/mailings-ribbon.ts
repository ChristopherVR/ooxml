import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';
import { group, row, stack, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';

export type MailingsCommand =
	| 'start'
	| 'source'
	| 'recipients'
	| 'highlight'
	| 'address'
	| 'greeting'
	| 'field'
	| 'rules'
	| 'match'
	| 'preview'
	| 'first'
	| 'previous'
	| 'next'
	| 'last'
	| 'find'
	| 'check'
	| 'finish';

export function buildMailings(panel: HTMLElement): void {
	const command = (key: MailingsCommand, label: string, icon: RibbonIcon, inline = false) => {
		const navigation = ['first', 'previous', 'next', 'last'].includes(key);
		const button = tool(
			label,
			icon,
			{ type: 'mailings', key },
			navigation ? {} : { large: !inline, inline },
		);
		button.dataset.mailings = key;
		if (key === 'preview' || key === 'highlight') button.setAttribute('aria-pressed', 'false');
		else if (!navigation) button.setAttribute('aria-haspopup', 'dialog');
		return button;
	};
	const unavailable = (label: string, icon: RibbonIcon, inline = false) => {
		const button = tool(label, icon, null, { large: !inline, inline });
		button.dataset.localetitle = 'mailings.envelopesUnavailable';
		return button;
	};
	panel.append(
		group('Create', unavailable('Envelopes', 'envelope'), unavailable('Labels', 'table')),
		group(
			'Start Mail Merge',
			command('start', 'Start Mail Merge', 'envelope'),
			command('source', 'Select Recipients', 'select'),
			command('recipients', 'Edit Recipient List', 'table'),
		),
		group(
			'Write & Insert Fields',
			command('highlight', 'Highlight Merge Fields', 'highlight'),
			command('address', 'Address Block', 'textBox'),
			command('greeting', 'Greeting Line', 'textBox'),
			command('field', 'Insert Merge Field', 'field'),
			stack(
				command('rules', 'Rules', 'field', true),
				command('match', 'Match Fields', 'field', true),
				unavailable('Update Labels', 'updateTable', true),
			),
		),
		group(
			'Preview Results',
			command('preview', 'Preview Results', 'find'),
			stack(
				row(
					command('first', 'First recipient', 'firstRecord', true),
					command('previous', 'Previous recipient', 'previousRecord', true),
					command('next', 'Next recipient', 'nextRecord', true),
					command('last', 'Last recipient', 'lastRecord', true),
				),
				command('find', 'Find Recipient', 'find', true),
				command('check', 'Check for Errors', 'spelling', true),
			),
		),
		group('Finish', command('finish', 'Finish & Merge', 'envelope')),
	);
}

export interface MailingsState {
	count: number;
	record: number;
	hasFields: boolean;
	hasSource: boolean;
	preview: boolean;
	highlight: boolean;
	busy: boolean;
}

export function mailingsReason(
	core: EditorCore,
	key: MailingsCommand,
	state = core.mailings?.state,
): LocalizationKey | undefined {
	if (core.loading || state?.busy) return 'mailings.busy';
	if (!core.view || !core.host.canEditOutsideBody()) return 'mailings.editUnavailable';
	if (state?.preview && ['start', 'address', 'greeting', 'field', 'rules', 'match'].includes(key))
		return 'mailings.exitPreview';
	if (['address', 'greeting', 'field', 'rules'].includes(key)) {
		const { $from, $to } = core.view.state.selection;
		if (core.targetView() !== core.view || !$from.sameParent($to) || !$from.parent.inlineContent)
			return 'mailings.selectText';
	}
	if (['field', 'match', 'address', 'greeting'].includes(key) && !state?.hasFields)
		return 'mailings.chooseSource';
	if (key === 'recipients' && !state?.hasSource) return 'mailings.chooseSource';
	if (['preview', 'find', 'finish'].includes(key) && !state?.count) return 'mailings.noRecipients';
	if (key === 'first' || key === 'previous' || key === 'next' || key === 'last') {
		if (!state?.preview) return 'mailings.startPreview';
		if ((key === 'first' || key === 'previous') && state.record <= 1) return 'mailings.atFirst';
		if ((key === 'next' || key === 'last') && state.record >= state.count) return 'mailings.atLast';
	}
	return undefined;
}

export function syncMailingsRibbon(core: EditorCore): void {
	for (const button of core.shell.toolbar?.querySelectorAll<HTMLButtonElement>('[data-mailings]') ??
		[]) {
		const key = button.dataset.mailings as MailingsCommand;
		const reason = mailingsReason(core, key);
		button.disabled = reason !== undefined;
		const label = button.dataset.localearialabel ?? button.getAttribute('aria-label') ?? '';
		button.title = translate(core.locale, reason ?? (label as LocalizationKey));
		if (key === 'preview' || key === 'highlight')
			button.setAttribute('aria-pressed', String(Boolean(core.mailings?.state[key])));
	}
}
