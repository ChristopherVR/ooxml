/** Value controls of the ribbon: drop-down selects, editable combo boxes and galleries. */
import { setLargeCaption } from './caption.js';
import type { Command } from '../commands';
import type { EditorContext } from '../context';
import { el, tagCommand, tooltip, type RenderScope } from './controls';
import { ribbonIcon } from './icons';
import type { RibbonControl, RibbonOption } from './parts';
import { arrowNavigation, closeRibbonPopover, mountPopover, openList } from './popover';

type SelectControl = Extract<RibbonControl, { kind: 'select' }>;
type GalleryControl = Extract<RibbonControl, { kind: 'gallery' }>;

const optionsOf = (ctx: EditorContext, control: SelectControl): RibbonOption[] => {
	try {
		return typeof control.options === 'function' ? control.options(ctx) : control.options;
	} catch {
		return [];
	}
};
const valueOf = (ctx: EditorContext, command: Command): string => {
	try {
		return command.value?.(ctx) ?? '';
	} catch {
		return '';
	}
};

/** Font name lists draw each entry in its own face. */
const isFontList = (id: string) => /font(-name|-family)?$/i.test(id);

/** A native `<select>` (number format, ...) or, with `editable`, a combo box (font, size). */
export function renderSelect(
	scope: RenderScope,
	control: SelectControl,
	command: Command,
): HTMLElement {
	return control.editable
		? renderCombo(scope, control, command)
		: renderNativeSelect(scope, control, command);
}

function renderNativeSelect(
	scope: RenderScope,
	control: SelectControl,
	command: Command,
): HTMLElement {
	const { ctx, doc } = scope;
	const select = el(doc, 'select', 'ribbon-select');
	tagCommand(select, command);
	select.setAttribute('aria-label', ctx.t(command.label));
	select.title = tooltip(ctx, command);
	if (control.width) select.style.width = `${control.width}px`;
	let signature = '';
	const fill = () => {
		const options = optionsOf(ctx, control);
		const next = JSON.stringify(options);
		if (next === signature) return;
		signature = next;
		select.replaceChildren(...options.map(({ value, label }) => new Option(ctx.t(label), value)));
	};
	select.addEventListener('change', () => void ctx.commands.run(command.id, select.value));
	scope.updates.push(() => {
		fill();
		select.disabled = !ctx.commands.isEnabled(command.id);
		const value = valueOf(ctx, command);
		let custom = select.querySelector<HTMLOptionElement>('option[data-custom]');
		const known = [...select.options].some(
			(option) => option.value === value && !option.dataset.custom,
		);
		if (!known && value) {
			custom ??= select.appendChild(new Option(value, value));
			custom.dataset.custom = '';
			custom.value = value;
			custom.textContent = value;
		} else custom?.remove();
		select.value = value;
		select.title = select.selectedOptions[0]?.textContent || tooltip(ctx, command);
	});
	return select;
}

function renderCombo(scope: RenderScope, control: SelectControl, command: Command): HTMLElement {
	const { ctx, doc } = scope;
	const wrap = el(doc, 'div', 'ribbon-combo');
	if (control.width) wrap.style.width = `${control.width}px`;
	const input = el(doc, 'input');
	input.type = 'text';
	tagCommand(input, command);
	input.setAttribute('role', 'combobox');
	input.setAttribute('aria-label', ctx.t(command.label));
	input.setAttribute('aria-autocomplete', 'list');
	input.setAttribute('aria-expanded', 'false');
	input.title = tooltip(ctx, command);
	input.autocomplete = 'off';
	input.spellcheck = false;
	let committed = '';
	const commit = (text: string) => {
		const value = text.trim();
		if (!value) {
			input.value = committed;
			return;
		}
		committed = value;
		input.value = value;
		void ctx.commands.run(command.id, value);
	};
	const caret = el(doc, 'button', 'ribbon-caret');
	caret.type = 'button';
	caret.tabIndex = -1;
	caret.setAttribute('aria-hidden', 'true');
	caret.append(ribbonIcon(doc, 'caret', 12));
	caret.addEventListener('mousedown', (event) => event.preventDefault());
	const open = () =>
		openList(
			wrap,
			optionsOf(ctx, control).map(({ value, label }) => ({ value, label: ctx.t(label) })),
			input.value,
			(value) => commit(value),
			{ preview: isFontList(command.id), expanded: input },
		);
	caret.addEventListener('click', open);
	input.addEventListener('focus', () => input.select());
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			commit(input.value);
			ctx.grid()?.focus();
		} else if (event.key === 'Escape') {
			input.value = committed;
			ctx.grid()?.focus();
		} else if (event.key === 'ArrowDown' && event.altKey === false) {
			event.preventDefault();
			open();
		}
	});
	input.addEventListener('change', () => commit(input.value));
	scope.updates.push(() => {
		const enabled = ctx.commands.isEnabled(command.id);
		input.disabled = !enabled;
		caret.disabled = !enabled;
		if (doc.activeElement === input || input.matches(':focus')) return;
		committed = valueOf(ctx, command);
		input.value = committed;
	});
	wrap.append(input, caret);
	return wrap;
}

/**
 * Parses preview SVG markup into a node without scripts, foreign content, event handlers or
 * script URLs (style names in a preview can come from the workbook).
 */
export function safeSvg(doc: Document, markup: string): Element | null {
	if (typeof DOMParser === 'undefined') return null;
	const parsed = new DOMParser().parseFromString(markup, 'image/svg+xml');
	const svg = parsed.documentElement;
	if (svg.nodeName !== 'svg' || parsed.getElementsByTagName('parsererror').length) return null;
	for (const node of [svg, ...svg.querySelectorAll('*')]) {
		if (/^(script|foreignObject|iframe|object|embed)$/i.test(node.localName)) {
			node.remove();
			continue;
		}
		for (const attribute of [...node.attributes]) {
			const value = attribute.value.trim().toLowerCase();
			if (
				/^on/i.test(attribute.name) ||
				value.startsWith('javascript:') ||
				(/href$/i.test(attribute.name) && !value.startsWith('#'))
			)
				node.removeAttribute(attribute.name);
		}
	}
	return doc.importNode(svg, true);
}

/** Applies a gallery preview: SVG markup (sanitized) or inline CSS around the label. */
function applyPreview(node: HTMLElement, preview: string | undefined, label: string): void {
	const svg = preview?.trimStart().startsWith('<svg') ? safeSvg(node.ownerDocument, preview) : null;
	if (svg) node.append(svg);
	else {
		if (preview && !preview.trimStart().startsWith('<')) node.style.cssText = preview;
		node.textContent = label;
	}
}

/**
 * A gallery command (Format as Table, Cell Styles): a large button with a caret, as Excel shows it
 * on a ribbon of ordinary width, opening every item as preview tiles in a grid.
 */
export function renderGallery(
	scope: RenderScope,
	control: GalleryControl,
	command: Command,
): HTMLElement {
	const { ctx, doc } = scope;
	const label = ctx.t(command.label);
	const button = el(doc, 'button', 'ribbon-large ribbon-gallery-button');
	button.type = 'button';
	tagCommand(button, command);
	button.setAttribute('aria-label', label);
	button.setAttribute('aria-haspopup', 'true');
	button.setAttribute('aria-expanded', 'false');
	button.title = tooltip(ctx, command);
	const caption = el(doc, 'span');
	setLargeCaption(caption, label);
	button.append(ribbonIcon(doc, command.icon, 28), caption, ribbonIcon(doc, 'caret', 12));
	button.addEventListener('mousedown', (event) => event.preventDefault());
	const items = () => {
		try {
			return control.items(ctx);
		} catch {
			return [];
		}
	};
	button.addEventListener('click', () => {
		const pop = el(doc, 'div', 'ribbon-popover ribbon-gallery-menu');
		pop.setAttribute('role', 'menu');
		pop.setAttribute('aria-label', label);
		for (const item of items()) {
			const tile = el(doc, 'button', 'ribbon-gallery-item');
			tile.type = 'button';
			tile.setAttribute('role', 'menuitem');
			const name = ctx.t(item.label);
			tile.setAttribute('aria-label', name);
			tile.title = name;
			const preview = el(doc, 'span', 'ribbon-gallery-preview');
			applyPreview(preview, item.preview, name);
			tile.append(preview);
			tile.addEventListener('mousedown', (event) => event.preventDefault());
			tile.addEventListener('click', () => {
				closeRibbonPopover();
				void ctx.commands.run(command.id, item.id);
			});
			pop.append(tile);
		}
		arrowNavigation(pop, '[role="menuitem"]');
		if (mountPopover(button, pop)) pop.querySelector<HTMLElement>('[role="menuitem"]')?.focus();
	});
	scope.updates.push(() => {
		button.disabled = !ctx.commands.isEnabled(command.id);
	});
	return button;
}
