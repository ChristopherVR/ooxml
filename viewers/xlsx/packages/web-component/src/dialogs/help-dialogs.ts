// Help tab dialogs: the keyboard shortcut list (fallback when the shell has no help command) and
// an honest About and Feature Status page.
import type { EditorContext } from '../context.js';
import { shortcutRows } from '../shortcut-help.js';
import { el, fieldset, text } from './fields.js';
import { showDialog } from './frame.js';

export function openShortcutHelp(ctx: EditorContext): Promise<undefined> {
	const doc = ctx.host.ownerDocument;
	const box = el(ctx, 'div', 'xve-results');
	const table = doc.createElement('table');
	const head = table.createTHead().insertRow();
	for (const label of ['Action', 'Keys']) {
		const th = doc.createElement('th');
		th.scope = 'col';
		th.textContent = ctx.t(label);
		head.append(th);
	}
	const body = table.createTBody();
	for (const r of shortcutRows(ctx)) {
		const tr = body.insertRow();
		tr.insertCell().textContent = r.label;
		tr.insertCell().textContent = r.keys;
	}
	box.append(table);
	box.style.maxHeight = '360px';
	return showDialog<undefined>(ctx, {
		name: 'shortcut-help',
		heading: 'Keyboard Shortcuts',
		okLabel: null,
		body: box,
		wide: true,
	});
}

export const SUPPORTED_FEATURES = [
	'Cell editing with undo and redo',
	'Formulas calculated by the built-in calculation engine',
	'Number formats',
	'Conditional formatting display',
	'Tables',
	'Charts (display and basic editing)',
	'Comments and notes',
	'Hyperlinks',
	'Data validation',
	'Sort and filter',
	'Freeze panes',
] as const;

export const UNSUPPORTED_FEATURES = [
	'Pivot tables and sparklines (kept in the file but not shown)',
	'Show Formulas',
	'Manual calculation',
	'Outline (grouping) for columns',
	'Password protection hashes',
	'Editing threaded comment replies',
	'Macros (VBA is kept in the file but never run)',
	'Changing the workbook theme',
	'Page Layout view',
] as const;

function bullets(ctx: EditorContext, items: readonly string[]): HTMLUListElement {
	const list = el(ctx, 'ul');
	for (const item of items) {
		const li = el(ctx, 'li');
		li.textContent = ctx.t(item);
		list.append(li);
	}
	return list;
}

export function openFeatureStatus(ctx: EditorContext): Promise<undefined> {
	return showDialog<undefined>(ctx, {
		name: 'feature-status',
		heading: 'About and Feature Status',
		okLabel: null,
		wide: true,
		body: [
			text(
				ctx,
				'This spreadsheet editor opens, edits and saves Excel workbooks in the browser. It is not Microsoft Excel and does not support every Excel feature.',
				'xve-message',
			),
			fieldset(ctx, 'Supported', bullets(ctx, SUPPORTED_FEATURES)),
			fieldset(ctx, 'Not supported yet', bullets(ctx, UNSUPPORTED_FEATURES)),
		],
	});
}
