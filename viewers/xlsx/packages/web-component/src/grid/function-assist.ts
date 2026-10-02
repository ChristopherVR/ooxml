// Function autocomplete list and argument tooltip, shared by the in-cell editor and the formula
// bar. Pure text logic lives in formula-text.ts; this module owns the popup DOM and its keys.
import { FUNCTION_CATALOG, type FunctionInfo } from '@christophervr/xlsx-core';
import { ensureStyle, h } from './dom.js';
import type { FormulaField } from './formula-field.js';
import {
	activeParam,
	callAtCaret,
	functionMatches,
	nameAtCaret,
	splice,
	syntaxParams,
} from './formula-text.js';

export interface FunctionAssist {
	update(field: FormulaField, anchor: { left: number; top: number }): void;
	/** True when the key was consumed by the popup. */
	handleKey(event: KeyboardEvent, field: FormulaField): boolean;
	close(): void;
	isOpen(): boolean;
	destroy(): void;
}

type Translate = (key: string, vars?: Record<string, string | number>) => string;

const CSS = `
.xg-assist{position:absolute;z-index:60;font:12px/1.4 'Segoe UI',system-ui,sans-serif;color:var(--xve-foreground,#252423)}
.xg-assist[hidden]{display:none}
.xg-assist-list{background:var(--xve-popover,#fff);border:1px solid var(--xve-border,#c8c6c4);box-shadow:0 4px 12px rgba(0,0,0,.16);min-width:180px;max-height:240px;overflow:auto;margin:0;padding:2px 0;list-style:none}
.xg-assist-option{padding:2px 8px;cursor:default;white-space:nowrap}
.xg-assist-option[aria-selected="true"]{background:var(--xve-accent,#e1dfdd)}
.xg-assist-desc{background:var(--xve-popover,#fff);border:1px solid var(--xve-border,#c8c6c4);border-top:0;padding:4px 8px;max-width:320px;white-space:normal;color:var(--xve-muted-foreground,#605e5c)}
.xg-assist-tip{background:var(--xve-popover,#fff);border:1px solid var(--xve-border,#c8c6c4);box-shadow:0 2px 6px rgba(0,0,0,.12);padding:2px 6px;white-space:nowrap}
.xg-assist-name{color:var(--xve-primary,#0f6cbd);text-decoration:underline}
.xg-assist-param-active{font-weight:700}
`;

const styleRoot = (container: HTMLElement): ShadowRoot | HTMLElement => {
	const root = container.getRootNode();
	return typeof ShadowRoot !== 'undefined' && root instanceof ShadowRoot ? root : container;
};

export function createFunctionAssist(
	doc: Document,
	container: HTMLElement,
	t: Translate,
	catalog: readonly FunctionInfo[] = FUNCTION_CATALOG,
): FunctionAssist {
	ensureStyle(styleRoot(container), 'xg-assist', CSS);
	const root = h(doc, 'div', 'xg-assist');
	root.hidden = true;
	const list = h(doc, 'ul', 'xg-assist-list', { role: 'listbox', 'aria-label': t('Functions') });
	const desc = h(doc, 'div', 'xg-assist-desc');
	const tip = h(doc, 'div', 'xg-assist-tip', { role: 'tooltip' });
	container.append(root);

	let matches: FunctionInfo[] = [];
	let selected = 0;
	let token: { start: number; prefix: string } | undefined;
	let lastField: FormulaField | undefined;

	const close = () => {
		matches = [];
		token = undefined;
		root.hidden = true;
		root.replaceChildren();
	};

	const renderList = () => {
		list.replaceChildren();
		matches.forEach((info, index) => {
			const option = h(doc, 'li', 'xg-assist-option', {
				role: 'option',
				'aria-selected': String(index === selected),
			});
			option.dataset.name = info.name;
			option.textContent = info.name;
			option.addEventListener('mousedown', (event) => {
				event.preventDefault();
				selected = index;
				if (lastField) accept(lastField);
			});
			list.append(option);
		});
		desc.textContent = matches[selected]?.description ?? '';
		root.replaceChildren(list, desc);
		list.children[selected]?.scrollIntoView?.({ block: 'nearest' });
	};

	const renderTip = (info: FunctionInfo, argIndex: number) => {
		const { name, params } = syntaxParams(info.syntax);
		const active = activeParam(params, argIndex);
		tip.replaceChildren();
		const nameSpan = h(doc, 'span', 'xg-assist-name');
		nameSpan.textContent = name || info.name;
		tip.append(nameSpan, doc.createTextNode('('));
		params.forEach((param, index) => {
			if (index > 0) tip.append(doc.createTextNode(', '));
			const span = h(
				doc,
				'span',
				index === active ? 'xg-assist-param xg-assist-param-active' : 'xg-assist-param',
			);
			span.textContent = param;
			tip.append(span);
		});
		tip.append(doc.createTextNode(')'));
		root.replaceChildren(tip);
	};

	const accept = (field: FormulaField) => {
		const info = matches[selected];
		if (!info || !token) return;
		const next = splice(
			field.input.value,
			token.start,
			field.input.selectionEnd ?? field.input.value.length,
			`${info.name}(`,
		);
		close();
		field.setText(next.text, next.caret);
		field.input.dispatchEvent(new Event('input', { bubbles: true }));
	};

	return {
		update(field, anchor) {
			lastField = field;
			const text = field.input.value;
			const caret = field.input.selectionEnd ?? text.length;
			root.style.left = `${anchor.left}px`;
			root.style.top = `${anchor.top}px`;
			token = nameAtCaret(text, caret);
			const found = token ? functionMatches(token.prefix, 12, catalog) : [];
			if (found.length) {
				const previous = matches[selected]?.name;
				matches = found;
				const keep = previous ? found.findIndex((m) => m.name === previous) : -1;
				selected = keep >= 0 ? keep : 0;
				root.hidden = false;
				renderList();
				return;
			}
			matches = [];
			const call = callAtCaret(text, caret);
			const info = call ? catalog.find((entry) => entry.name === call.name) : undefined;
			if (call && info) {
				root.hidden = false;
				renderTip(info, call.argIndex);
				return;
			}
			close();
		},
		handleKey(event, field) {
			if (root.hidden) return false;
			if (!matches.length) {
				if (event.key === 'Escape') {
					close();
					return true;
				}
				return false;
			}
			if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
				const step = event.key === 'ArrowDown' ? 1 : -1;
				selected = (selected + step + matches.length) % matches.length;
				renderList();
			} else if (event.key === 'Tab' || (event.key === 'Enter' && !event.altKey)) {
				accept(field);
			} else if (event.key === 'Escape') {
				close();
			} else return false;
			event.preventDefault();
			event.stopPropagation();
			return true;
		},
		close,
		isOpen: () => !root.hidden,
		destroy() {
			close();
			root.remove();
		},
	};
}
