/**
 * "Tell me what you want to do": searches every registered, enabled, visible command by its
 * translated and English label and runs the chosen one. A command that needs a value (font, size,
 * number format, colour) takes the user to its ribbon control instead.
 */
import type { Command } from '../commands';
import type { EditorContext } from '../context';
import { el } from './controls';
import { ribbonIcon } from './icons';

export interface TellMeHandlers {
	isHidden(id: string): boolean;
	/** Shows the ribbon tab holding a command's control and focuses it; false when there is none. */
	revealControl(id: string): boolean;
}

const MAX_RESULTS = 8;

/** Commands matching `query` (case-insensitive, any word order), best first. */
export function searchCommands(
	ctx: EditorContext,
	query: string,
	isHidden: (id: string) => boolean,
): Command[] {
	const words = query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
	if (!words.length) return [];
	const scored: Array<{ command: Command; score: number }> = [];
	const seen = new Set<string>();
	for (const command of ctx.commands.list()) {
		if (isHidden(command.id) || !ctx.commands.isEnabled(command.id)) continue;
		const label = ctx.t(command.label).toLocaleLowerCase();
		const english = command.label.toLocaleLowerCase();
		const haystack = `${label} ${english} ${command.id.replace(/[.-]/g, ' ')}`;
		if (!words.every((word) => haystack.includes(word))) continue;
		if (seen.has(label)) continue;
		seen.add(label);
		const score = label.startsWith(words[0]!) ? 0 : label.includes(words[0]!) ? 1 : 2;
		scored.push({ command, score });
	}
	return scored
		.sort(
			(a, b) => a.score - b.score || ctx.t(a.command.label).localeCompare(ctx.t(b.command.label)),
		)
		.slice(0, MAX_RESULTS)
		.map(({ command }) => command);
}

export interface TellMe {
	element: HTMLElement;
	focus(): void;
	relocalize(): void;
}

export function createTellMe(ctx: EditorContext, handlers: TellMeHandlers): TellMe {
	const doc = ctx.host.ownerDocument;
	const box = el(doc, 'div', 'xve-tellme');
	box.append(ribbonIcon(doc, 'find', 16));
	const input = el(doc, 'input');
	input.type = 'search';
	input.setAttribute('role', 'combobox');
	input.setAttribute('aria-expanded', 'false');
	input.setAttribute('aria-controls', 'xve-tellme-results');
	input.setAttribute('aria-autocomplete', 'list');
	const list = el(doc, 'ul', 'xve-tellme-results');
	list.id = 'xve-tellme-results';
	list.setAttribute('role', 'listbox');
	list.hidden = true;
	let matches: Command[] = [];
	let active = 0;
	const close = () => {
		list.hidden = true;
		input.setAttribute('aria-expanded', 'false');
	};
	const run = (index: number) => {
		const command = matches[index];
		if (!command) return;
		close();
		input.value = '';
		const needsValue = Boolean(command.value);
		if (needsValue && handlers.revealControl(command.id)) return;
		ctx.grid()?.focus();
		void ctx.commands.run(command.id);
	};
	const mark = () => {
		list
			.querySelectorAll('[role="option"]')
			.forEach((item, index) => item.setAttribute('aria-selected', String(index === active)));
		input.setAttribute(
			'aria-activedescendant',
			matches.length ? `xve-tellme-option-${active}` : '',
		);
	};
	const render = () => {
		matches = searchCommands(ctx, input.value, handlers.isHidden);
		active = 0;
		list.replaceChildren(
			...matches.map((command, index) => {
				const item = el(doc, 'li');
				item.id = `xve-tellme-option-${index}`;
				item.setAttribute('role', 'option');
				item.append(ribbonIcon(doc, command.icon, 16));
				const text = el(doc, 'span');
				text.textContent = ctx.t(command.label);
				item.append(text);
				if (command.shortcut) {
					const hint = el(doc, 'span', 'xve-tellme-hint');
					hint.textContent = command.shortcut;
					item.append(hint);
				}
				item.addEventListener('mousedown', (event) => event.preventDefault());
				item.addEventListener('click', () => run(index));
				return item;
			}),
		);
		const query = input.value.trim();
		if (query && !matches.length) {
			const empty = el(doc, 'li', 'xve-tellme-empty');
			empty.textContent = ctx.t('No matching commands');
			list.append(empty);
		}
		list.hidden = !query;
		input.setAttribute('aria-expanded', String(!list.hidden));
		mark();
	};
	input.addEventListener('input', render);
	input.addEventListener('blur', close);
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			input.value = '';
			close();
			ctx.grid()?.focus();
		} else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			if (!matches.length) return;
			event.preventDefault();
			active = (active + (event.key === 'ArrowDown' ? 1 : matches.length - 1)) % matches.length;
			mark();
		} else if (event.key === 'Enter') {
			event.preventDefault();
			run(active);
		}
	});
	const relocalize = () => {
		const text = ctx.t('Tell me what you want to do');
		input.placeholder = text;
		input.setAttribute('aria-label', text);
	};
	relocalize();
	box.append(input, list);
	return { element: box, focus: () => input.focus(), relocalize };
}
