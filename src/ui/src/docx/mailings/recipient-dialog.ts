import type { MailMergeData, MailMergeRecord } from 'ooxml-core/docx/mail-merge';
import { checkbox, dialogButton, labelled, row, selectOf, textInput } from '../dialog-fields';
import type { EditorLocale } from '../localization';
import { dialogSubmit, type MailingsDialog } from './dialog';
import { recipientText, selectedRecipients, type RecipientSelection } from './recipient-data';

const PAGE_SIZE = 20;

export function editRecipientList(
	dialog: MailingsDialog,
	locale: EditorLocale,
	data: MailMergeData,
	initial: RecipientSelection,
	apply: (settings: RecipientSelection) => void,
): void {
	const { t } = dialog;
	const fields = (data.fields ?? []).map((field) => [field, field] as const);
	const filter = textInput();
	filter.maxLength = 128;
	filter.value = initial.filter;
	const filterBy = selectOf([['', t('All columns')], ...fields]);
	filterBy.value = initial.filterField;
	const sort = selectOf([['', t('Source order')], ...fields]);
	sort.value = initial.sortField;
	const descending = checkbox(t('Descending'));
	descending.input.checked = initial.descending;
	const wrapper = document.createElement('div');
	wrapper.className = 'dve-mailings-table';
	const table = document.createElement('table');
	table.setAttribute('aria-label', dialog.element.dataset.title ?? '');
	wrapper.append(table);
	const header = table.createTHead().insertRow();
	for (const label of [t('Include'), t('Record'), ...fields.map(([name]) => name)]) {
		const th = document.createElement('th');
		th.scope = 'col';
		th.textContent = label;
		header.append(th);
	}
	const body = table.createTBody();
	const count = document.createElement('output');
	count.setAttribute('aria-live', 'polite');
	const previous = dialogButton(t('Previous'));
	const next = dialogButton(t('Next'));
	const pager = document.createElement('div');
	pager.className = 'dve-mailings-pager';
	pager.append(previous, count, next);
	dialog.content.append(
		row(labelled(t('Filter recipients'), filter), labelled(t('Filter column'), filterBy)),
		row(labelled(t('Sort by'), sort), descending.wrapper),
		wrapper,
		pager,
	);
	const excluded = new Set(initial.excluded);
	const settings = (): RecipientSelection => ({
		excluded,
		filter: filter.value,
		filterField: filterBy.value,
		sortField: sort.value,
		descending: descending.input.checked,
	});
	const sourceIndex = new Map(data.records.map((record, index) => [record, index + 1]));
	let visible: readonly MailMergeRecord[] = [];
	let page = 0;
	const render = () => {
		const start = page * PAGE_SIZE;
		const end = Math.min(start + PAGE_SIZE, visible.length);
		body.replaceChildren();
		for (const record of visible.slice(start, end)) {
			const tr = body.insertRow();
			const number = String(sourceIndex.get(record));
			const include = document.createElement('input');
			include.type = 'checkbox';
			include.checked = !excluded.has(record);
			include.setAttribute('aria-label', `${t('Include')}: ${t('Record')} ${number}`);
			include.addEventListener('change', () => {
				if (include.checked) excluded.delete(record);
				else excluded.add(record);
			});
			tr.insertCell().append(include);
			tr.insertCell().textContent = number;
			for (const [field] of fields) {
				const cell = tr.insertCell();
				cell.textContent = recipientText(record, field).slice(0, 256);
			}
		}
		count.value = `${visible.length ? start + 1 : 0}–${end} / ${visible.length}`;
		previous.disabled = page === 0;
		next.disabled = end === visible.length;
	};
	const update = () => {
		visible = selectedRecipients(data, settings(), locale, true).records;
		page = 0;
		render();
	};
	let timer: ReturnType<typeof setTimeout> | undefined;
	filter.addEventListener('input', () => {
		clearTimeout(timer);
		timer = setTimeout(() => {
			if (dialog.element.isConnected) update();
		}, 120);
	});
	for (const control of [filterBy, sort, descending.input])
		control.addEventListener('change', update);
	previous.addEventListener('click', () => {
		page--;
		render();
	});
	next.addEventListener('click', () => {
		page++;
		render();
	});
	dialogSubmit(dialog, locale, () => {
		clearTimeout(timer);
		dialog.close();
		apply(settings());
	});
	update();
}
