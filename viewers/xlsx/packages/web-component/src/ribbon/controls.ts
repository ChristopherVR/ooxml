/**
 * Renders one declarative ribbon control (see parts.ts) into DOM. Every control reads its label,
 * icon, tooltip and state from its command, and registers an `update` hook the ribbon calls on
 * `requestRender` to refresh enabled, pressed and value states.
 */
import { setLargeCaption } from './caption.js';
import type { Color } from '@christophervr/xlsx-core';
import type { Command } from '../commands';
import type { EditorContext } from '../context';
import { cssColor, openColorGrid } from './color-grid';
import { renderSelect, renderGallery } from './controls-input';
import { ribbonIcon } from './icons';
import { isMenuSeparator, type RibbonControl, type RibbonMenuItem } from './parts';
import { openMenu, type MenuEntry } from './popover';

export interface RenderScope {
	ctx: EditorContext;
	doc: Document;
	/** Refresh hooks, run on every `requestRender`. */
	updates: Array<() => void>;
	/** Command ids the host hid (`hiddenActions`). */
	isHidden(id: string): boolean;
}

export const el = <K extends keyof HTMLElementTagNameMap>(
	doc: Document,
	tag: K,
	className = '',
): HTMLElementTagNameMap[K] => {
	const node = doc.createElement(tag);
	if (className) node.className = className;
	return node;
};

/** `Bold (Ctrl+B)` tooltip text for a command. */
export const tooltip = (ctx: EditorContext, command: Command): string =>
	command.shortcut ? `${ctx.t(command.label)} (${command.shortcut})` : ctx.t(command.label);

/** Marks a control with its command id (hidden actions, Tell me, KeyTips, Customize Ribbon). */
export function tagCommand(node: HTMLElement, command: Command): void {
	node.dataset.command = command.id;
	node.dataset.labelKey = command.label;
}

function plainButton(scope: RenderScope, className: string): HTMLButtonElement {
	const button = el(scope.doc, 'button', className);
	button.type = 'button';
	// Keep keyboard focus (and the in-cell editor) where it was, as Office does.
	button.addEventListener('mousedown', (event) => event.preventDefault());
	return button;
}

/** A command button: `large` stacks a 28px icon over its caption; small shows a 16px icon. */
export function commandButton(
	scope: RenderScope,
	command: Command,
	options: {
		size?: 'large' | 'small';
		showLabel?: boolean;
		toggle?: boolean;
		caption?: boolean;
	} = {},
): HTMLButtonElement {
	const { ctx, doc } = scope;
	const large = options.size === 'large';
	const button = plainButton(scope, large ? 'ribbon-large' : 'ribbon-small');
	tagCommand(button, command);
	const label = ctx.t(command.label);
	button.setAttribute('aria-label', label);
	button.title = tooltip(ctx, command);
	button.append(ribbonIcon(doc, command.icon, large ? 28 : 16));
	if ((large && options.caption !== false) || options.showLabel || !command.icon) {
		const caption = el(doc, 'span');
		if (large) setLargeCaption(caption, label);
		else caption.textContent = label;
		button.append(caption);
		if (!large) button.classList.add('ribbon-labelled');
	}
	button.addEventListener('click', () => void ctx.commands.run(command.id));
	scope.updates.push(() => {
		button.disabled = !ctx.commands.isEnabled(command.id);
		if (options.toggle)
			button.setAttribute('aria-pressed', String(safe(() => command.checked?.(ctx)) ?? false));
	});
	return button;
}

const safe = <T>(read: () => T): T | undefined => {
	try {
		return read();
	} catch {
		return undefined;
	}
};

/** Menu entries for spec items; items whose command is not registered are left out. */
export function menuEntries(scope: RenderScope, items: RibbonMenuItem[]): MenuEntry[] {
	const { ctx } = scope;
	const entries: MenuEntry[] = [];
	for (const item of items) {
		if (isMenuSeparator(item)) {
			if (entries.length && !('separator' in entries[entries.length - 1]!)) entries.push(item);
			continue;
		}
		const command = ctx.commands.get(item.command);
		if (!command || scope.isHidden(command.id)) continue;
		const checked = item.arg === undefined ? safe(() => command.checked?.(ctx)) : undefined;
		entries.push({
			label: ctx.t(item.label ?? command.label),
			...(command.icon ? { icon: command.icon } : {}),
			...(command.shortcut ? { shortcut: command.shortcut } : {}),
			...(checked === undefined ? {} : { checked }),
			disabled: !ctx.commands.isEnabled(command.id),
			run: () => void ctx.commands.run(command.id, item.arg),
		});
	}
	while (entries.length && 'separator' in entries[entries.length - 1]!) entries.pop();
	return entries;
}

function caretButton(scope: RenderScope, label: string): HTMLButtonElement {
	const caret = plainButton(scope, 'ribbon-caret');
	caret.dataset.splitCaret = '';
	caret.setAttribute('aria-haspopup', 'true');
	caret.setAttribute('aria-expanded', 'false');
	caret.setAttribute('aria-label', scope.ctx.t('{label} options', { label }));
	caret.title = caret.getAttribute('aria-label') ?? '';
	caret.append(ribbonIcon(scope.doc, 'caret', 12));
	return caret;
}

function renderSplit(
	scope: RenderScope,
	control: Extract<RibbonControl, { kind: 'split' }>,
	command: Command,
) {
	const { ctx, doc } = scope;
	const large = control.size === 'large';
	const wrap = el(doc, 'div', large ? 'ribbon-split ribbon-split-large' : 'ribbon-split');
	const main = commandButton(scope, command, { size: control.size ?? 'small', caption: false });
	const label = ctx.t(command.label);
	const caret = caretButton(scope, label);
	if (large) {
		const caption = el(doc, 'span');
		setLargeCaption(caption, label);
		caret.prepend(caption);
	}
	caret.addEventListener('click', () => openMenu(caret, menuEntries(scope, control.menu), label));
	scope.updates.push(() => {
		caret.disabled =
			main.disabled &&
			menuEntries(scope, control.menu).every((e) => 'separator' in e || e.disabled);
	});
	wrap.append(main, caret);
	return wrap;
}

function renderMenu(scope: RenderScope, control: Extract<RibbonControl, { kind: 'menu' }>) {
	const { ctx, doc } = scope;
	const large = control.size === 'large';
	const label = ctx.t(control.label);
	const button = plainButton(
		scope,
		large ? 'ribbon-large ribbon-menu-button' : 'ribbon-small ribbon-labelled ribbon-menu-button',
	);
	button.dataset.labelKey = control.label;
	button.dataset.menuCommands = control.items
		.flatMap((item) => (isMenuSeparator(item) ? [] : [item.command]))
		.join(' ');
	button.setAttribute('aria-label', label);
	button.setAttribute('aria-haspopup', 'true');
	button.setAttribute('aria-expanded', 'false');
	button.title = label;
	button.append(ribbonIcon(doc, control.icon, large ? 28 : 16));
	const caption = el(doc, 'span');
	if (large) setLargeCaption(caption, label);
	else caption.textContent = label;
	button.append(caption, ribbonIcon(doc, 'caret', 12));
	button.addEventListener('click', () =>
		openMenu(button, menuEntries(scope, control.items), label),
	);
	scope.updates.push(() => {
		button.disabled = menuEntries(scope, control.items).every(
			(e) => 'separator' in e || e.disabled,
		);
	});
	return button;
}

/** Default colour of a colour split before the user picks one (Excel: red font, yellow fill). */
const initialColor = (id: string): Color | undefined =>
	/fill|highlight/i.test(id) ? { rgb: 'FFFF00' } : /font/i.test(id) ? { rgb: 'FF0000' } : undefined;

function renderColor(
	scope: RenderScope,
	control: Extract<RibbonControl, { kind: 'color' }>,
	command: Command,
) {
	const { ctx, doc } = scope;
	const wrap = el(doc, 'div', 'ribbon-split');
	const label = ctx.t(control.label);
	const main = plainButton(scope, 'ribbon-small ribbon-color');
	tagCommand(main, command);
	main.setAttribute('aria-label', label);
	main.title = command.shortcut ? `${label} (${command.shortcut})` : label;
	main.append(ribbonIcon(doc, control.icon, 16));
	let current = initialColor(command.id);
	const paint = () => main.style.setProperty('--bar', cssColor(current, ctx.workbook()?.theme));
	paint();
	main.addEventListener('click', () => void ctx.commands.run(command.id, current));
	const caret = caretButton(scope, label);
	const fill = /fill|highlight/i.test(command.id);
	caret.addEventListener('click', () =>
		openColorGrid(
			caret,
			(color) => {
				current = color;
				paint();
				void ctx.commands.run(command.id, color);
			},
			{
				t: ctx.t,
				automaticLabel: fill ? 'No Fill' : 'Automatic',
				...(ctx.workbook() ? { theme: ctx.workbook()!.theme } : {}),
			},
		),
	);
	scope.updates.push(() => {
		const enabled = ctx.commands.isEnabled(command.id);
		main.disabled = !enabled;
		caret.disabled = !enabled;
		paint();
	});
	wrap.append(main, caret);
	return wrap;
}

/** Renders a control, or null when its command is not registered. */
export function renderControl(
	scope: RenderScope,
	control: RibbonControl,
	inStack = false,
): HTMLElement | null {
	const { ctx, doc } = scope;
	switch (control.kind) {
		case 'separator':
			return el(doc, 'div', 'ribbon-separator');
		case 'stack': {
			const stack = el(doc, 'div', 'ribbon-stack');
			for (const child of control.controls) {
				const node = renderControl(scope, child, true);
				if (node) stack.append(node);
			}
			return stack.childElementCount ? stack : null;
		}
		case 'menu':
			return renderMenu(scope, control);
	}
	const command = ctx.commands.get(control.command);
	if (!command) return null;
	switch (control.kind) {
		case 'button':
			return commandButton(scope, command, {
				size: inStack ? 'small' : (control.size ?? 'small'),
				...(control.showLabel === undefined ? {} : { showLabel: control.showLabel }),
			});
		case 'toggle':
			return commandButton(scope, command, {
				size: inStack ? 'small' : (control.size ?? 'small'),
				toggle: true,
				...(control.showLabel === undefined ? {} : { showLabel: control.showLabel }),
			});
		case 'split':
			return renderSplit(scope, control, command);
		case 'color':
			return renderColor(scope, control, command);
		case 'select':
			return renderSelect(scope, control, command);
		case 'gallery':
			return renderGallery(scope, control, command);
	}
}
