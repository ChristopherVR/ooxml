import {
	mailMergeAddressBlock,
	mailMergeAddressFields,
	mailMergeComparisons,
	mailMergeField,
	mailMergeGreetingLine,
	mailMergeRule,
	matchMailMergeField,
	setMailMergeFieldMapping,
	type MailMergeFieldMapping,
} from 'ooxml-core/docx/mail-merge';
import type { TextRun } from 'ooxml-core/docx';
import { checkbox, labelled, row, selectOf, textInput } from '../dialog-fields';
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
	CourtesyTitle: 'Courtesy title',
	MiddleName: 'Middle name',
	Suffix: 'Name suffix',
	Nickname: 'Nickname',
	SpouseCourtesyTitle: 'Partner courtesy title',
	SpouseFirstName: 'Partner first name',
	SpouseMiddleName: 'Partner middle name',
	SpouseLastName: 'Partner last name',
	SpouseNickname: 'Partner nickname',
	EmailAddress: 'Email address column',
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

export function nativeFieldDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	kind: 'address' | 'greeting',
	fields: readonly string[],
	mappings: readonly MailMergeFieldMapping[],
	insert: (runs: TextRun[]) => void,
): void {
	const { t } = dialog;
	const name = selectOf([
		['full', t('First and last name')],
		['first', t('First name')],
		['last', t('Title and last name')],
		...(kind === 'address' ? ([['none', t('Omit name')]] as const) : []),
	]);
	if (!matchMailMergeField('LastName', fields, mappings)) name.value = 'first';
	dialog.content.append(labelled(t('Recipient name'), name));
	const nameValue = () => (name.value === 'first' || name.value === 'last' ? name.value : 'full');
	let runs: () => TextRun[];
	if (kind === 'greeting') {
		const greeting = textInput();
		const punctuation = textInput();
		const fallback = textInput();
		greeting.maxLength = fallback.maxLength = 128;
		punctuation.maxLength = 8;
		greeting.value = t('Dear');
		punctuation.value = ',';
		fallback.value = t('To whom it may concern,');
		dialog.content.append(
			labelled(t('Greeting'), greeting),
			labelled(t('Punctuation'), punctuation),
			labelled(t('Greeting for missing names'), fallback),
		);
		runs = () =>
			mailMergeGreetingLine({
				name: nameValue(),
				before: greeting.value,
				after: punctuation.value,
				fallback: fallback.value,
			});
	} else {
		const company = checkbox(t('Include company'));
		const street = checkbox(t('Include postal address'));
		const order = checkbox(t('Format address for recipient country'));
		company.input.checked = street.input.checked = order.input.checked = true;
		const country = selectOf([
			['1', t('Always include country')],
			['0', t('Omit country')],
			['2', t('Exclude named country')],
		]);
		const exclude = textInput();
		exclude.maxLength = 128;
		const exclusion = labelled(t('Country or region'), exclude);
		exclusion.hidden = true;
		country.addEventListener('change', () => (exclusion.hidden = country.value !== '2'));
		dialog.content.append(
			company.wrapper,
			street.wrapper,
			order.wrapper,
			labelled(t('Country or region'), country),
			exclusion,
		);
		runs = () =>
			mailMergeAddressBlock({
				name: name.value === 'none' ? name.value : nameValue(),
				company: company.input.checked,
				street: street.input.checked,
				formatByCountry: order.input.checked,
				country: country.value === '0' || country.value === '2' ? country.value : '1',
				...(country.value === '2' && { excludeCountry: exclude.value }),
			});
	}
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
	const once = checkbox(t('Ask once for this merge'));
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
	dialog.content.append(
		labelled(t('Rule'), kind),
		condition,
		branches,
		bookmarkRow,
		prompts,
		once.wrapper,
	);
	const missing = dialogMessage(dialog, 'Select recipients before adding a comparison rule.');
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
					once: once.input.checked,
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
		once.wrapper.hidden = !usesPrompt;
		ok.disabled =
			!missing.hidden || (named && !/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(bookmark.value));
	};
	kind.addEventListener('change', update);
	bookmark.addEventListener('input', update);
	update();
}
