// Small DOM builders for dialog forms. Every label is an English key translated through ctx.t.
import type { EditorContext } from '../context.js';

type El = HTMLElement;

const make = <K extends keyof HTMLElementTagNameMap>(
	ctx: EditorContext,
	tag: K,
	className?: string,
): HTMLElementTagNameMap[K] => {
	const el = ctx.host.ownerDocument.createElement(tag);
	if (className) el.className = className;
	return el;
};

export const el = make;

let uid = 0;
export const nextId = (prefix: string): string => `xve-${prefix}-${++uid}`;

/** A labelled control in one row: `<label>text</label><control>`. */
export function field<T extends El>(
	ctx: EditorContext,
	label: string,
	control: T,
): HTMLLabelElement {
	const wrapper = make(ctx, 'label', 'xve-field');
	const text = make(ctx, 'span', 'xve-field-label');
	text.textContent = ctx.t(label);
	control.setAttribute('aria-label', ctx.t(label));
	wrapper.append(text, control);
	return wrapper;
}

export function textInput(ctx: EditorContext, value = '', name?: string): HTMLInputElement {
	const input = make(ctx, 'input', 'xve-input');
	input.type = 'text';
	input.value = value;
	if (name) input.name = name;
	return input;
}

export function numberInput(
	ctx: EditorContext,
	value: number,
	min: number,
	max: number,
	step = 1,
): HTMLInputElement {
	const input = make(ctx, 'input', 'xve-input xve-input-number');
	input.type = 'number';
	input.min = String(min);
	input.max = String(max);
	input.step = String(step);
	input.value = String(value);
	return input;
}

export function textArea(ctx: EditorContext, value = '', rows = 3): HTMLTextAreaElement {
	const area = make(ctx, 'textarea', 'xve-input');
	area.rows = rows;
	area.value = value;
	return area;
}

/** A native select; option labels are English keys unless `raw` is set. */
export function select(
	ctx: EditorContext,
	options: ReadonlyArray<readonly [value: string, label: string]>,
	value?: string,
	raw = false,
): HTMLSelectElement {
	const sel = make(ctx, 'select', 'xve-input');
	for (const [v, label] of options) {
		const option = make(ctx, 'option');
		option.value = v;
		option.textContent = raw ? label : ctx.t(label);
		sel.append(option);
	}
	if (value !== undefined) sel.value = value;
	return sel;
}

export function checkbox(
	ctx: EditorContext,
	label: string,
	checked = false,
): { wrapper: HTMLLabelElement; input: HTMLInputElement } {
	const wrapper = make(ctx, 'label', 'xve-check');
	const input = make(ctx, 'input');
	input.type = 'checkbox';
	input.checked = checked;
	input.setAttribute('aria-label', ctx.t(label));
	const text = make(ctx, 'span');
	text.textContent = ctx.t(label);
	wrapper.append(input, text);
	return { wrapper, input };
}

/** A radio group; returns the wrapper and a getter / setter for the chosen value. */
export function radios(
	ctx: EditorContext,
	legend: string,
	options: ReadonlyArray<readonly [value: string, label: string]>,
	value: string,
): {
	element: HTMLFieldSetElement;
	get(): string;
	set(v: string): void;
	inputs: HTMLInputElement[];
} {
	const name = nextId('radio');
	const inputs: HTMLInputElement[] = [];
	const set = make(ctx, 'fieldset', 'xve-fieldset');
	const title = make(ctx, 'legend');
	title.textContent = ctx.t(legend);
	set.append(title);
	for (const [v, label] of options) {
		const wrapper = make(ctx, 'label', 'xve-check');
		const input = make(ctx, 'input');
		input.type = 'radio';
		input.name = name;
		input.value = v;
		input.checked = v === value;
		input.setAttribute('aria-label', ctx.t(label));
		const text = make(ctx, 'span');
		text.textContent = ctx.t(label);
		wrapper.append(input, text);
		set.append(wrapper);
		inputs.push(input);
	}
	return {
		element: set,
		inputs,
		get: () => inputs.find((i) => i.checked)?.value ?? value,
		set(v) {
			for (const input of inputs) input.checked = input.value === v;
		},
	};
}

export function fieldset(
	ctx: EditorContext,
	legend: string,
	...children: El[]
): HTMLFieldSetElement {
	const set = make(ctx, 'fieldset', 'xve-fieldset');
	const title = make(ctx, 'legend');
	title.textContent = ctx.t(legend);
	set.append(title, ...children);
	return set;
}

export function row(ctx: EditorContext, ...children: El[]): HTMLDivElement {
	const div = make(ctx, 'div', 'xve-row');
	div.append(...children);
	return div;
}

export function text(
	ctx: EditorContext,
	key: string,
	className = 'xve-note',
	vars?: Record<string, string | number>,
): HTMLParagraphElement {
	const p = make(ctx, 'p', className);
	p.textContent = ctx.t(key, vars);
	return p;
}

/** A list box (`role="listbox"`) of options with keyboard selection; `raw` labels are shown as is. */
export function listBox(
	ctx: EditorContext,
	label: string,
	onChange?: (value: string) => void,
	onActivate?: (value: string) => void,
): {
	element: HTMLDivElement;
	setItems(items: ReadonlyArray<readonly [value: string, label: string]>, raw?: boolean): void;
	value(): string | undefined;
	select(value: string): void;
} {
	const box = make(ctx, 'div', 'xve-listbox');
	box.setAttribute('role', 'listbox');
	box.setAttribute('aria-label', ctx.t(label));
	box.tabIndex = 0;
	let selected: string | undefined;
	const options = (): HTMLElement[] => [...box.querySelectorAll<HTMLElement>('[role="option"]')];
	const choose = (value: string, notify = true): void => {
		selected = value;
		for (const option of options()) {
			const on = option.dataset.value === value;
			option.setAttribute('aria-selected', String(on));
			if (on) {
				box.setAttribute('aria-activedescendant', option.id);
				option.scrollIntoView?.({ block: 'nearest' });
			}
		}
		if (notify) onChange?.(value);
	};
	box.addEventListener('keydown', (event) => {
		const list = options();
		const index = list.findIndex((o) => o.dataset.value === selected);
		const move = (to: number): void => {
			const next = list[Math.max(0, Math.min(list.length - 1, to))];
			if (next?.dataset.value !== undefined) choose(next.dataset.value);
		};
		if (event.key === 'ArrowDown') move(index + 1);
		else if (event.key === 'ArrowUp') move(index - 1);
		else if (event.key === 'Home') move(0);
		else if (event.key === 'End') move(list.length - 1);
		else if (event.key === 'Enter' && selected !== undefined && onActivate) onActivate(selected);
		else return;
		event.preventDefault();
		event.stopPropagation();
	});
	return {
		element: box,
		setItems(items, raw = true) {
			box.replaceChildren(
				...items.map(([value, text]) => {
					const option = make(ctx, 'div', 'xve-option');
					option.setAttribute('role', 'option');
					option.id = nextId('opt');
					option.dataset.value = value;
					option.textContent = raw ? text : ctx.t(text);
					option.addEventListener('click', () => choose(value));
					option.addEventListener('dblclick', () => onActivate?.(value));
					return option;
				}),
			);
			if (selected !== undefined && items.some(([v]) => v === selected)) choose(selected, false);
			else selected = undefined;
		},
		value: () => selected,
		select: (value) => choose(value),
	};
}

/** Accessible tabs (`role="tablist"`, arrow keys move between tabs). */
export function tabs(
	ctx: EditorContext,
	pages: ReadonlyArray<{ id: string; label: string; panel: HTMLElement }>,
	initial?: string,
): { element: HTMLDivElement; show(id: string): void; current(): string } {
	const wrapper = make(ctx, 'div', 'xve-tabs');
	const list = make(ctx, 'div', 'xve-tablist');
	list.setAttribute('role', 'tablist');
	const buttons: HTMLButtonElement[] = [];
	let current = initial ?? pages[0]?.id ?? '';
	const show = (id: string): void => {
		current = id;
		pages.forEach((page, i) => {
			const on = page.id === id;
			const b = buttons[i];
			if (b) {
				b.setAttribute('aria-selected', String(on));
				b.tabIndex = on ? 0 : -1;
			}
			page.panel.hidden = !on;
		});
	};
	pages.forEach((page, i) => {
		const b = make(ctx, 'button', 'xve-tab');
		b.type = 'button';
		b.id = nextId('tab');
		b.setAttribute('role', 'tab');
		b.dataset.tab = page.id;
		b.textContent = ctx.t(page.label);
		page.panel.setAttribute('role', 'tabpanel');
		page.panel.setAttribute('aria-labelledby', b.id);
		page.panel.classList.add('xve-tabpanel');
		b.addEventListener('click', () => show(page.id));
		b.addEventListener('keydown', (event) => {
			const delta = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
			if (!delta) return;
			event.preventDefault();
			const next = pages[(i + delta + pages.length) % pages.length];
			if (!next) return;
			show(next.id);
			buttons[pages.indexOf(next)]?.focus();
		});
		buttons.push(b);
		list.append(b);
	});
	wrapper.append(list, ...pages.map((p) => p.panel));
	show(current);
	return { element: wrapper, show, current: () => current };
}

export function panel(ctx: EditorContext, ...children: El[]): HTMLDivElement {
	const div = make(ctx, 'div');
	div.append(...children);
	return div;
}

/** Marks a field invalid with a translated message, focusing it; returns undefined. */
export function invalid(
	ctx: EditorContext,
	control: HTMLElement,
	message: string,
	vars?: Record<string, string | number>,
): undefined {
	control.setAttribute('aria-invalid', 'true');
	ctx.toast(ctx.t(message, vars), 'warning');
	control.focus();
	return undefined;
}
