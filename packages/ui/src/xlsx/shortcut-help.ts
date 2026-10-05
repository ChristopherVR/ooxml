/** The keyboard shortcut help dialog (Ctrl+/ or F1): the global map plus every command's shortcut. */
import type { EditorContext } from './context';
import { displayKeys, isMacPlatform, SHORTCUTS } from './keyboard';
import { el } from './ribbon/controls';

export interface ShortcutRow {
	label: string;
	keys: string;
}

/** Rows in display order: the global shortcuts, then commands that declare their own (grid keys). */
export function shortcutRows(ctx: EditorContext, mac = isMacPlatform()): ShortcutRow[] {
	const rows: ShortcutRow[] = [];
	const seen = new Set<string>();
	for (const shortcut of SHORTCUTS) {
		const command = shortcut.command ? ctx.commands.get(shortcut.command) : undefined;
		if (shortcut.command && !command) continue;
		const keys = displayKeys(shortcut, mac).join(' / ');
		rows.push({ label: ctx.t(command?.label ?? shortcut.label), keys });
		if (shortcut.command) seen.add(shortcut.command);
	}
	for (const command of ctx.commands.list()) {
		if (!command.shortcut || seen.has(command.id)) continue;
		rows.push({ label: ctx.t(command.label), keys: command.shortcut });
	}
	// Several commands can share a label and keys (the ribbon's and the grid's Delete).
	const unique = new Set<string>();
	return rows.filter((row) => {
		const key = `${row.label}|${row.keys}`;
		if (unique.has(key)) return false;
		unique.add(key);
		return true;
	});
}

export interface ShortcutHelp {
	readonly element: HTMLElement;
	readonly isOpen: boolean;
	open(): void;
	close(): void;
}

export function createShortcutHelp(ctx: EditorContext, onClose: () => void): ShortcutHelp {
	const doc = ctx.host.ownerDocument;
	const element = el(doc, 'section', 'xve-dialog xve-shortcut-help');
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-modal', 'true');
	element.hidden = true;
	const heading = el(doc, 'h2');
	heading.id = 'xve-shortcut-help-title';
	element.setAttribute('aria-labelledby', heading.id);
	const note = el(doc, 'p');
	const table = el(doc, 'table');
	const closeButton = el(doc, 'button', 'xve-dialog-primary');
	closeButton.type = 'button';
	const actions = el(doc, 'div', 'xve-dialog-actions');
	actions.append(closeButton);
	element.append(heading, note, table, actions);
	const close = () => {
		if (element.hidden) return;
		element.hidden = true;
		onClose();
	};
	closeButton.addEventListener('click', close);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			close();
		} else if (event.key === 'Tab') {
			event.preventDefault();
			closeButton.focus();
		}
	});
	return {
		element,
		get isOpen() {
			return !element.hidden;
		},
		open() {
			heading.textContent = ctx.t('Keyboard shortcuts');
			note.textContent = ctx.t('Grid keys (arrows, Enter, Tab, F2) work while a cell is selected.');
			closeButton.textContent = ctx.t('Close');
			table.replaceChildren();
			const head = table.createTHead().insertRow();
			for (const text of ['Action', 'Keys']) {
				const cell = el(doc, 'th');
				cell.scope = 'col';
				cell.textContent = ctx.t(text);
				head.append(cell);
			}
			const body = table.createTBody();
			for (const row of shortcutRows(ctx)) {
				const tr = body.insertRow();
				tr.insertCell().textContent = row.label;
				const kbd = el(doc, 'kbd');
				kbd.textContent = row.keys;
				tr.insertCell().append(kbd);
			}
			element.hidden = false;
			element.scrollTop = 0;
			closeButton.focus({ preventScroll: true });
		},
		close,
	};
}
