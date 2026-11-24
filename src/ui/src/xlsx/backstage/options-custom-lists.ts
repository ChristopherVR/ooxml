/** File > Options > Advanced > Edit Custom Lists: add, edit, delete and import sort lists. */
import { BUILT_IN_SORT_LISTS, formatRange, parseRange, parseSortList } from 'ooxml-core/xlsx';
import {
	deleteUserSortList,
	importSortListFromRange,
	saveUserSortList,
	updateUserSortList,
	userSortLists,
} from '../dialogs/custom-lists';
import type { BackstageHost } from './parts';

const NEW_KEY = 'new';
const builtinKey = (i: number): string => `builtin:${i}`;
const userKey = (i: number): string => `user:${i}`;

/**
 * Renders Excel's Edit Custom Lists manager inside the Advanced options pane (`pane-advanced`).
 * Built-in lists are read-only; user lists can be added, edited, deleted or imported from cells.
 */
export function renderCustomListsPane(pane: HTMLElement, host: BackstageHost): void {
	const { ctx } = host;
	const { t } = ctx;
	const doc = pane.ownerDocument;
	let selected = NEW_KEY;

	const heading = doc.createElement('h3');
	heading.textContent = t('Edit Custom Lists');

	const note = doc.createElement('p');
	note.className = 'xve-options-note';
	note.textContent = t('Create lists for use in sorts and fill sequences.');

	const layout = doc.createElement('div');
	layout.className = 'xve-custom-lists';

	const listCol = doc.createElement('div');
	listCol.className = 'xve-custom-lists-col';
	const listLabel = doc.createElement('label');
	listLabel.textContent = t('Custom lists:');
	const listBox = doc.createElement('select');
	listBox.className = 'xve-custom-lists-select';
	listBox.size = 6;
	listBox.setAttribute('aria-label', t('Custom lists:'));
	listCol.append(listLabel, listBox);

	const entriesCol = doc.createElement('div');
	entriesCol.className = 'xve-custom-lists-col';
	const entriesLabel = doc.createElement('label');
	entriesLabel.textContent = t('List entries:');
	const entries = doc.createElement('textarea');
	entries.className = 'xve-custom-lists-entries';
	entries.rows = 6;
	entries.setAttribute('aria-label', t('List entries:'));
	entries.placeholder = t('Press Enter to separate list entries.');
	entriesCol.append(entriesLabel, entries);

	const actionsCol = doc.createElement('div');
	actionsCol.className = 'xve-custom-lists-actions';
	const saveBtn = doc.createElement('button');
	saveBtn.type = 'button';
	saveBtn.className = 'xve-options-action';
	const deleteBtn = doc.createElement('button');
	deleteBtn.type = 'button';
	deleteBtn.className = 'xve-options-action';
	deleteBtn.textContent = t('Delete');
	actionsCol.append(saveBtn, deleteBtn);

	layout.append(listCol, entriesCol, actionsCol);

	// Import list from cells
	const importRow = doc.createElement('div');
	importRow.className = 'xve-custom-lists-import';
	const importLabel = doc.createElement('label');
	importLabel.textContent = t('Import list from cells:');
	const importInput = doc.createElement('input');
	importInput.type = 'text';
	importInput.className = 'xve-custom-lists-range';
	importInput.setAttribute('aria-label', t('Import list from cells:'));
	const activeSel = ctx.selection.get().ranges[0];
	importInput.value = activeSel ? formatRange(activeSel) : 'A1:A5';
	const importBtn = doc.createElement('button');
	importBtn.type = 'button';
	importBtn.className = 'xve-options-action';
	importBtn.textContent = t('Import');
	importRow.append(importLabel, importInput, importBtn);

	const status = doc.createElement('p');
	status.className = 'xve-options-note';
	status.setAttribute('role', 'status');

	const syncEditor = () => {
		const users = userSortLists();
		if (selected === NEW_KEY) {
			entries.value = '';
			entries.readOnly = false;
			saveBtn.textContent = t('Add');
			saveBtn.disabled = false;
			deleteBtn.disabled = true;
			return;
		}
		if (selected.startsWith('builtin:')) {
			const idx = Number(selected.slice('builtin:'.length));
			entries.value = (BUILT_IN_SORT_LISTS[idx] ?? []).join('\n');
			entries.readOnly = true;
			saveBtn.textContent = t('Add');
			saveBtn.disabled = true;
			deleteBtn.disabled = true;
			return;
		}
		if (selected.startsWith('user:')) {
			const idx = Number(selected.slice('user:'.length));
			entries.value = (users[idx] ?? []).join('\n');
			entries.readOnly = false;
			saveBtn.textContent = t('Save');
			saveBtn.disabled = false;
			deleteBtn.disabled = false;
		}
	};

	const rebuildList = () => {
		const users = userSortLists();
		listBox.replaceChildren();
		const addOpt = (value: string, label: string) => {
			const opt = doc.createElement('option');
			opt.value = value;
			opt.textContent = label;
			listBox.append(opt);
		};
		addOpt(NEW_KEY, t('NEW LIST'));
		BUILT_IN_SORT_LISTS.forEach((items, i) => addOpt(builtinKey(i), items.join(', ')));
		users.forEach((items, i) => addOpt(userKey(i), items.join(', ')));
		if (selected.startsWith('user:')) {
			const idx = Number(selected.slice('user:'.length));
			if (idx < 0 || idx >= users.length) selected = NEW_KEY;
		}
		listBox.value = selected;
		syncEditor();
	};

	listBox.addEventListener('change', () => {
		selected = listBox.value;
		status.textContent = '';
		syncEditor();
	});

	saveBtn.addEventListener('click', () => {
		const items = parseSortList(entries.value);
		if (items.length === 0) {
			status.textContent = t('Enter at least one list entry.');
			return;
		}
		if (selected === NEW_KEY) {
			const idx = saveUserSortList(items);
			if (idx >= 0) selected = userKey(idx);
		} else if (selected.startsWith('user:')) {
			const idx = Number(selected.slice('user:'.length));
			updateUserSortList(idx, items);
		}
		status.textContent = '';
		rebuildList();
	});

	deleteBtn.addEventListener('click', () => {
		if (!selected.startsWith('user:')) return;
		const idx = Number(selected.slice('user:'.length));
		if (deleteUserSortList(idx)) {
			selected = NEW_KEY;
			status.textContent = '';
			rebuildList();
		}
	});

	importBtn.addEventListener('click', () => {
		const raw = importInput.value.trim().replace(/\$/g, '');
		const range = parseRange(raw);
		const sheet = ctx.workbook()?.sheets[ctx.activeSheet()];
		if (!range || !sheet) {
			status.textContent = t('Enter a valid cell range such as A1:A5.');
			return;
		}
		const items = importSortListFromRange(sheet, range);
		if (items.length === 0) {
			status.textContent = t('The selected range has no text or number entries to import.');
			return;
		}
		const idx = saveUserSortList(items);
		if (idx >= 0) {
			selected = userKey(idx);
			status.textContent = '';
			rebuildList();
		}
	});

	rebuildList();
	pane.replaceChildren(heading, note, layout, importRow, status);
}
