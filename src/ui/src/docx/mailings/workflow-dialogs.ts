import type {
	MailMergeData,
	MailMergeDataSource,
	MailMergeIssue,
} from 'ooxml-core/docx/mail-merge';
import { dialogButton, labelled, selectOf, textInput } from '../dialog-fields';
import type { EditorLocale } from '../localization';
import { dialogChoice, dialogMessage, dialogSubmit, type MailingsDialog } from './dialog';
import { recipientMatches } from './recipient-data';
import type { MailingsLabel } from './strings';

export type MergeType = 'formLetters' | 'catalog' | 'envelopes' | 'mailingLabels';

export function startDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	current: string | undefined,
	blank: boolean,
	apply: (type: MergeType | undefined) => void,
): void {
	const select = selectOf([
		['formLetters', dialog.t('Letters')],
		['catalog', dialog.t('Directory')],
		['envelopes', dialog.t('Envelopes')],
		['mailingLabels', dialog.t('Labels')],
		['', dialog.t('Normal Word Document')],
	]);
	if (!blank) {
		for (const option of select.options)
			if (option.value === 'envelopes' || option.value === 'mailingLabels')
				option.disabled = option.value !== current;
		dialogMessage(dialog, 'Start with an empty document to set up envelopes or labels.');
	}
	select.value = current ?? '';
	if (select.selectedIndex < 0) select.value = '';
	dialog.content.append(labelled(dialog.element.dataset.title ?? '', select));
	dialogSubmit(dialog, locale, () => {
		const type = select.value;
		dialog.close();
		apply(
			type === 'catalog' ||
				type === 'formLetters' ||
				type === 'envelopes' ||
				type === 'mailingLabels'
				? type
				: undefined,
		);
	});
}

export function sourceDialog(
	dialog: MailingsDialog,
	provider: MailMergeDataSource | undefined,
	choose: (file?: File) => void,
): void {
	const input = document.createElement('input');
	input.type = 'file';
	input.accept = '.csv,.tsv,.xlsx';
	input.hidden = true;
	input.addEventListener('change', () => {
		const file = input.files?.[0];
		if (!file) return;
		dialog.close();
		choose(file);
	});
	dialog.content.append(
		dialogChoice(dialog, 'Choose a CSV, TSV, or XLSX file', () => input.click()),
		input,
		dialogChoice(
			dialog,
			'Use host recipients',
			() => {
				dialog.close();
				choose();
			},
			provider ? undefined : 'The host has not registered a recipient provider.',
		),
	);
	dialog.actions();
}

export function findDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	data: MailMergeData,
	record: number,
	found: (record: number) => void,
): void {
	const query = textInput();
	query.maxLength = 128;
	dialog.content.append(labelled(dialog.t('Search recipient values'), query));
	const status = document.createElement('p');
	status.setAttribute('role', 'status');
	dialog.content.append(status);
	const next = dialogButton(dialog.t('Find next'), true);
	const search = () => {
		const term = query.value.trim().toLocaleLowerCase(locale);
		if (!term) return;
		const size = data.records.length;
		for (let offset = 0; offset < size; offset++) {
			const index = (record + offset) % size;
			const candidate = data.records[index];
			if (candidate && recipientMatches(candidate, data.fields ?? [], term, locale)) {
				dialog.close();
				found(index + 1);
				return;
			}
		}
		status.textContent = dialog.t('No matching recipient.');
	};
	next.disabled = true;
	query.addEventListener('input', () => {
		status.textContent = '';
		next.disabled = !query.value.trim();
	});
	query.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			search();
		}
	});
	next.addEventListener('click', search);
	dialog.actions(next);
}

const issueLabels: Record<MailMergeIssue['code'], MailingsLabel> = {
	'no-recipients': 'No recipients are included.',
	'unknown-field': 'Unknown field',
	'invalid-instruction': 'Invalid field instruction',
	'unsupported-field': 'Unsupported merge field',
	'invalid-record-rule-location': 'Record rule in an unsupported location',
};

export function checkDialog(dialog: MailingsDialog, issues: readonly MailMergeIssue[]): void {
	if (!issues.length) dialogMessage(dialog, 'No merge errors found.');
	else {
		const list = document.createElement('ul');
		for (const issue of issues.slice(0, 100)) {
			const item = document.createElement('li');
			const label = dialog.t(issueLabels[issue.code]);
			item.textContent =
				issue.code === 'no-recipients'
					? label
					: `${dialog.t('Record')} ${issue.recordNumber} · ${label} · ${issue.field.slice(0, 128)}`;
			list.append(item);
		}
		dialog.content.append(list);
		if (issues.length > 100) dialogMessage(dialog, 'Additional issues were found.');
	}
	dialog.actions();
}

export type MergeDestination = 'document' | 'print' | 'email';

export function finishDialog(
	dialog: MailingsDialog,
	finish: (destination: MergeDestination) => void,
	emailReason: MailingsLabel | undefined,
): void {
	const choose = (destination: MergeDestination) => () => {
		dialog.close();
		finish(destination);
	};
	dialog.content.append(
		dialogChoice(dialog, 'Edit Individual Documents', choose('document')),
		dialogChoice(dialog, 'Print Documents', choose('print')),
		dialogChoice(dialog, 'Send Email Messages', choose('email'), emailReason),
	);
	dialog.actions();
}
