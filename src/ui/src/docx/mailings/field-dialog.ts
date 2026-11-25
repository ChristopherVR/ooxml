import {
	composeMailMergeFields,
	mailMergeAddressFields,
	mailMergeComparisons,
	mailMergeField,
	mailMergeRule,
	matchMailMergeField,
	setMailMergeFieldMapping,
	type MailMergeFieldMapping,
} from 'ooxml-core/docx/mail-merge';
import type { TextRun } from 'ooxml-core/docx';
import { labelled, row, selectOf, textInput } from '../dialog-fields';
import type { EditorLocale } from '../localization';
import { dialogMessage, dialogSubmit, type MailingsDialog } from './dialog';
import type { MailingsLabel } from './strings';

const fieldLabels: Record<(typeof mailMergeAddressFields)[number], MailingsLabel> = {
	FirstName: 'First name',
	LastName: 'Last name',
	Company: 'Company',
	Address1: 'Address line 1',
	Address2: 'Address line 2',
	City: 'City',
	State: 'State or province',
	PostalCode: 'Postal code',
	CountryorRegion: 'Country or region',
};

export function fieldDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	fields: readonly string[],
	insert: (runs: TextRun[]) => void,
): void {
	const select = selectOf(fields.map((field) => [field, field]));
	dialog.content.append(labelled(dialog.t('Merge field'), select));
	dialogSubmit(dialog, locale, () => {
		dialog.close();
		insert([mailMergeField(select.value)]);
	});
}

export function matchFieldsDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	fields: readonly string[],
	current: readonly MailMergeFieldMapping[],
	apply: (mappings: MailMergeFieldMapping[]) => void,
): void {
	const controls = mailMergeAddressFields.map((name) => {
		const select = selectOf([
			['', dialog.t('Unmatched')],
			...fields.map((field) => [field, field] as const),
		]);
		select.value = matchMailMergeField(name, fields, current);
		dialog.content.append(labelled(dialog.t(fieldLabels[name]), select));
		return { name, select };
	});
	dialogSubmit(dialog, locale, () => {
		// Keep the source positions of unrelated Word ODSO entries.
		const mappings = current.map((mapping) => ({ ...mapping }));
		for (const { name, select } of controls)
			setMailMergeFieldMapping(mappings, name, fields, select.value);
		dialog.close();
		apply(mappings);
	});
}

export function composedFieldDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	kind: 'address' | 'greeting',
	fields: readonly string[],
	mappings: readonly MailMergeFieldMapping[],
	insert: (runs: TextRun[]) => void,
): void {
	const greeting = textInput();
	greeting.maxLength = 128;
	greeting.value = dialog.t('Dear');
	const runs = () => composeMailMergeFields(kind, fields, mappings, greeting.value);
	const example = document.createElement('p');
	example.className = 'dve-mailings-example';
	const update = () =>
		(example.textContent = runs()
			.map((run) => run.text)
			.join(''));
	update();
	if (!example.textContent) {
		dialogMessage(dialog, 'Match a name or address column first.');
		dialog.actions();
		return;
	}
	if (kind === 'greeting') {
		dialog.content.append(labelled(dialog.t('Greeting'), greeting));
		greeting.addEventListener('input', update);
	}
	dialog.content.append(example);
	dialogSubmit(dialog, locale, () => {
		dialog.close();
		insert(runs());
	});
}

export function rulesDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	fields: readonly string[],
	insert: (runs: TextRun[]) => void,
): void {
	const { t } = dialog;
	const rules = [
		['IF', 'IF…THEN…ELSE'],
		['NEXT', 'NEXT'],
		['NEXTIF', 'NEXT IF'],
		['SKIPIF', 'SKIP IF'],
		['MERGEREC', 'MERGEREC'],
		['MERGESEQ', 'MERGESEQ'],
		['ASK', 'ASK'],
		['FILLIN', 'FILL-IN'],
		['REF', t('Insert default answer (REF)')],
	] as const;
	const kind = selectOf(rules);
	const field = selectOf(fields.map((name) => [name, name]));
	const comparison = selectOf(mailMergeComparisons.map((op) => [op, op]));
	const value = textInput();
	const yes = textInput();
	const no = textInput();
	const bookmark = textInput();
	const prompt = textInput();
	const answer = textInput();
	for (const input of [value, yes, no, prompt, answer]) input.maxLength = 1024;
	bookmark.maxLength = 40;
	const condition = row(
		labelled(t('Merge field'), field),
		labelled(t('Comparison'), comparison),
		labelled(t('Compare with'), value),
	);
	const branches = row(labelled(t('Then insert'), yes), labelled(t('Otherwise insert'), no));
	const bookmarkRow = labelled(t('Bookmark'), bookmark);
	const prompts = row(labelled(t('Prompt'), prompt), labelled(t('Default answer'), answer));
	dialog.content.append(labelled(t('Rule'), kind), condition, branches, bookmarkRow, prompts);
	const missing = dialogMessage(dialog, 'Select recipients before adding a comparison rule.');
	const defaultNote = dialogMessage(dialog, 'Rule prompts use the default answer during merge.');
	const ok = dialogSubmit(dialog, locale, () => {
		const op = mailMergeComparisons.find((item) => item === comparison.value);
		let runs: TextRun[];
		switch (kind.value) {
			case 'IF':
			case 'NEXTIF':
			case 'SKIPIF':
				if (!op || !field.value) return;
				runs = mailMergeRule({
					kind: kind.value,
					field: field.value,
					operator: op,
					value: value.value,
					yes: yes.value,
					no: no.value,
				});
				break;
			case 'ASK':
			case 'FILLIN':
				runs = mailMergeRule({
					kind: kind.value,
					bookmark: bookmark.value,
					prompt: prompt.value,
					answer: answer.value,
				});
				break;
			case 'REF':
				runs = mailMergeRule({ kind: 'REF', bookmark: bookmark.value });
				break;
			case 'NEXT':
			case 'MERGEREC':
			case 'MERGESEQ':
				runs = mailMergeRule({ kind: kind.value });
				break;
			default:
				return;
		}
		dialog.close();
		insert(runs);
	});
	const update = () => {
		const conditional = ['IF', 'NEXTIF', 'SKIPIF'].includes(kind.value);
		const named = kind.value === 'ASK' || kind.value === 'REF';
		const usesPrompt = kind.value === 'ASK' || kind.value === 'FILLIN';
		condition.hidden = !conditional || !field.value;
		branches.hidden = kind.value !== 'IF' || !field.value;
		bookmarkRow.hidden = !named;
		prompts.hidden = !usesPrompt;
		missing.hidden = !conditional || Boolean(field.value);
		defaultNote.hidden = !usesPrompt;
		ok.disabled =
			!missing.hidden || (named && !/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(bookmark.value));
	};
	kind.addEventListener('change', update);
	bookmark.addEventListener('input', update);
	update();
}
