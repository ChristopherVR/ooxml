// The Name Box: shows the active cell (or the defined name of the selection); typing a reference
// or name and pressing Enter goes there; an unused valid name defines it for the selection.
import { formatAddress, type CellRange } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { h } from '../grid/dom.js';
import { absoluteReference, nameForRange, resolveNameBox } from './name-resolve.js';

export interface NameBox {
	readonly root: HTMLElement;
	readonly input: HTMLInputElement;
	refresh(): void;
	destroy(): void;
}

/** Focus check that also works inside a shadow root. */
export const isFocused = (node: HTMLElement): boolean =>
	(node.getRootNode() as Document | ShadowRoot).activeElement === node;

export function createNameBox(ctx: EditorContext, doc: Document): NameBox {
	const t = ctx.t;
	const root = h(doc, 'div', 'xfb-namebox');
	const input = h(doc, 'input', 'xfb-name', {
		part: 'name-box',
		type: 'text',
		role: 'combobox',
		'aria-label': t('Name Box'),
		'aria-expanded': 'false',
		'aria-autocomplete': 'none',
		spellcheck: 'false',
	});
	const toggle = h(doc, 'button', 'xfb-name-toggle', {
		type: 'button',
		'aria-label': t('Defined names'),
		tabindex: '-1',
	});
	toggle.innerHTML =
		'<svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true"><path d="M2 3.5l3 3 3-3" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>';
	const list = h(doc, 'ul', 'xfb-name-list', { role: 'listbox', 'aria-label': t('Defined names') });
	list.hidden = true;
	root.append(input, toggle, list);

	const current = (): string => {
		const wb = ctx.workbook();
		const selection = ctx.selection.get();
		const only = selection.ranges.length === 1 ? selection.ranges[0] : undefined;
		if (wb && only) {
			const name = nameForRange(wb, selection.sheet, only);
			if (name) return name;
		}
		return formatAddress(selection.active);
	};
	const refresh = () => {
		if (isFocused(input)) return;
		input.value = current();
	};

	const go = (sheet: number, range: CellRange) => {
		if (sheet !== ctx.activeSheet()) ctx.setActiveSheet(sheet);
		const active = { ...range.start };
		ctx.selection.set({ sheet, active, anchor: { ...range.start }, ranges: [range] });
		ctx.grid()?.scrollTo(active);
		ctx.grid()?.focus();
	};

	const submit = () => {
		const wb = ctx.workbook();
		const session = ctx.session();
		if (!wb) return;
		const result = resolveNameBox(wb, ctx.activeSheet(), input.value);
		if (result.kind === 'go') {
			go(result.sheet, result.range);
		} else if (result.kind === 'define' && session && !ctx.readOnly()) {
			const selection = ctx.selection.get();
			const range = selection.ranges[0];
			const sheet = wb.sheets[selection.sheet];
			if (!range || !sheet) return;
			session.setDefinedName({ name: result.name, formula: absoluteReference(sheet.name, range) });
			input.value = result.name;
			ctx.grid()?.focus();
		} else {
			ctx.toast(t("Reference isn't valid."), 'warning');
			input.select();
			return;
		}
		input.value = current();
	};

	const closeList = () => {
		list.hidden = true;
		input.setAttribute('aria-expanded', 'false');
	};
	const openList = () => {
		const wb = ctx.workbook();
		list.replaceChildren();
		const sheet = ctx.activeSheet();
		const names = (wb?.definedNames ?? []).filter(
			(d) => !d.hidden && (d.localSheet === undefined || d.localSheet === sheet),
		);
		for (const named of names) {
			const item = h(doc, 'li', 'xfb-name-item', { role: 'option' });
			item.textContent = named.name;
			item.addEventListener('mousedown', (event) => {
				event.preventDefault();
				closeList();
				input.value = named.name;
				submit();
			});
			list.append(item);
		}
		if (!names.length) {
			const empty = h(doc, 'li', 'xfb-name-empty', { 'aria-disabled': 'true' });
			empty.textContent = t('No defined names');
			list.append(empty);
		}
		list.hidden = false;
		input.setAttribute('aria-expanded', 'true');
	};

	const onKey = (event: KeyboardEvent) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			closeList();
			submit();
		} else if (event.key === 'Escape') {
			event.preventDefault();
			closeList();
			input.value = current();
			ctx.grid()?.focus();
		} else if (event.key === 'ArrowDown' && event.altKey) {
			event.preventDefault();
			openList();
		}
		event.stopPropagation();
	};
	const onFocus = () => input.select();
	const onBlur = () => {
		closeList();
		input.value = current();
	};
	const onToggle = (event: Event) => {
		event.preventDefault();
		if (list.hidden) openList();
		else closeList();
		input.focus();
	};
	input.addEventListener('keydown', onKey);
	input.addEventListener('focus', onFocus);
	input.addEventListener('blur', onBlur);
	toggle.addEventListener('mousedown', onToggle);
	refresh();

	return {
		root,
		input,
		refresh,
		destroy() {
			input.removeEventListener('keydown', onKey);
			input.removeEventListener('focus', onFocus);
			input.removeEventListener('blur', onBlur);
			toggle.removeEventListener('mousedown', onToggle);
			root.remove();
		},
	};
}
