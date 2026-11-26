// The Equation tab's galleries: the built-in equations, the symbols by category and the eleven
// structure galleries, each drawn as the equation it inserts (empty arguments as placeholders).
// A gallery is a dialog holding one or more labelled menus; the arrow keys move between its items
// (one tab stop, roving), Escape returns the caret to the equation.
import { EQUATION_GALLERY, EQUATION_PLACEHOLDER } from 'ooxml-core/math';
import { equationKit as kit } from './equation-kit';
import { equationPreview, renderEquation } from '../math/equation-render';
import type { RibbonAction } from './ribbon-action';
import { equationStyle } from './equation-editing';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import { localeOf, translate, type LocalizationKey } from './localization';

export interface GalleryHandlers {
	choose(action: RibbonAction): void;
	/** Escape closed the gallery without a choice. */
	dismissed?(): void;
}

type Say = (key: string) => string;
let menus = 0;

/** A labelled menu of the gallery, with a visible heading when `heading` is set. */
function menu(pop: HTMLElement, label: string, heading = false): HTMLElement {
	const list = document.createElement('div');
	list.setAttribute('role', 'menu');
	list.className = 'dve-equation-gallery-menu';
	if (heading) {
		const title = document.createElement('div');
		title.className = 'dve-equation-gallery-heading';
		title.id = `dve-equation-menu-${++menus}`;
		title.textContent = label;
		list.setAttribute('aria-labelledby', title.id);
		pop.append(title);
	} else list.setAttribute('aria-label', label);
	pop.append(list);
	return list;
}

function item(list: HTMLElement, label: string, choose: () => void): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.setAttribute('role', 'menuitem');
	button.setAttribute('aria-label', label);
	button.tabIndex = -1;
	button.title = label;
	// The equation editor keeps its caret while a gallery item is pressed.
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', () => {
		closeRibbonPopover();
		choose();
	});
	list.append(button);
	return button;
}

function preview(button: HTMLElement, latex: string, label: string): void {
	const omml = kit().buildUpEquation(latex, 'latex', false);
	const drawn = omml ? equationPreview(omml, true) : undefined;
	const math = drawn && renderEquation(button.ownerDocument, drawn, false, label);
	if (math) button.append(math);
	else button.textContent = label;
}

/** Arrow keys, Home and End move between the items of all the gallery's menus: one tab stop. */
function roving(pop: HTMLElement): void {
	const items = () => [...pop.querySelectorAll<HTMLElement>('[role="menuitem"]')];
	const focus = (target: HTMLElement | undefined) => {
		if (!target) return;
		for (const other of items()) other.tabIndex = other === target ? 0 : -1;
		target.focus({ preventScroll: false });
	};
	pop.addEventListener('keydown', (event) => {
		const all = items();
		const at = all.indexOf(event.target as HTMLElement);
		if (at < 0) return;
		const step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[event.key];
		if (step) focus(all[(at + step + all.length) % all.length]);
		else if (event.key === 'Home') focus(all[0]);
		else if (event.key === 'End') focus(all.at(-1));
		else return;
		event.preventDefault();
	});
	pop.addEventListener('focusin', (event) => {
		if (items().includes(event.target as HTMLElement)) focus(event.target as HTMLElement);
	});
}

function structures(pop: HTMLElement, group: string, say: Say, choose: GalleryHandlers['choose']) {
	const gallery = kit().EQUATION_STRUCTURES.find((entry) => entry.id === group);
	if (!gallery) return;
	// Word's galleries list the empty structures first, then common filled-in ones.
	const sections = [
		[say(gallery.name), gallery.structures.filter((s) => s.latex.includes(EQUATION_PLACEHOLDER))],
		[
			say('Common Structures'),
			gallery.structures.filter((s) => !s.latex.includes(EQUATION_PLACEHOLDER)),
		],
	] as const;
	for (const [heading, entries] of sections) {
		if (!entries.length) continue;
		const list = menu(pop, heading, true);
		for (const structure of entries) {
			const label = say(structure.name);
			const value = `${group}/${structure.id}`;
			preview(
				item(list, label, () => choose({ type: 'equation', key: 'structure', value })),
				structure.latex,
				label,
			);
		}
	}
}

function builtIns(
	pop: HTMLElement,
	say: Say,
	choose: GalleryHandlers['choose'],
	fromInsert: boolean,
) {
	const list = menu(pop, say('Built-In Equations'), true);
	for (const entry of EQUATION_GALLERY) {
		const label = say(entry.name);
		const button = item(list, label, () =>
			choose({ type: 'equation', key: 'builtIn', value: entry.id }),
		);
		button.classList.add('dve-equation-gallery-wide');
		const name = document.createElement('span');
		name.className = 'gallery-name';
		name.textContent = label;
		button.append(name);
		preview(button, entry.latex, label);
	}
	const commands = menu(pop, say('Equation'));
	commands.classList.add('dve-equation-gallery-commands');
	const command = (key: string, action: RibbonAction) =>
		(item(commands, say(key), () => choose(action)).textContent = say(key));
	command('Insert New Equation', { type: 'equation', key: 'new' });
	if (fromInsert) command('Insert or Edit Equation…', { type: 'formatDialog', kind: 'equation' });
}

/** A symbol's accessible name: the character, its math AutoCorrect word if any, its code point. */
function symbolName(symbol: string, words: ReadonlyMap<string, string>): string {
	const code = symbol.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0');
	const word = words.get(symbol);
	return `${symbol}${word ? `, \\${word}` : ''}, U+${code}`;
}

function symbols(pop: HTMLElement, say: Say, choose: GalleryHandlers['choose']) {
	const words = new Map<string, string>();
	for (const [word, char] of Object.entries(kit().MATH_AUTOCORRECT))
		if (!words.has(char)) words.set(char, word);
	const category = document.createElement('select');
	category.setAttribute('aria-label', say('Symbol category'));
	category.className = 'dve-equation-symbol-category';
	for (const entry of kit().EQUATION_SYMBOLS) category.add(new Option(say(entry.name), entry.id));
	pop.append(category);
	const grid = menu(pop, say(kit().EQUATION_SYMBOLS[0]!.name));
	grid.classList.add('dve-equation-symbol-grid');
	const show = () => {
		grid.replaceChildren();
		const entries = kit().EQUATION_SYMBOLS;
		const entry = entries.find((c) => c.id === category.value) ?? entries[0]!;
		grid.setAttribute('aria-label', say(entry.name));
		for (const symbol of entry.symbols)
			// Text presentation, so arrows and similar characters never draw as emoji.
			item(grid, symbolName(symbol, words), () =>
				choose({ type: 'equation', key: 'symbol', value: symbol }),
			).append(`${symbol}︎`);
		const first = grid.querySelector<HTMLElement>('[role="menuitem"]');
		if (first) first.tabIndex = 0;
	};
	category.addEventListener('change', show);
	show();
}

/**
 * Opens gallery `kind` (`builtIn`, `symbols` or a structure group) under `anchor`. An anchor on
 * the Insert tab (`data-equation-gallery="insert-builtIn"`) also offers the LaTeX dialog.
 */
export function openEquationGallery(
	anchor: HTMLElement,
	kind: string,
	handlers: GalleryHandlers,
): void {
	const locale = localeOf(anchor);
	// Structure and symbol category names come with the galleries; the rest is the editor's.
	const names: Record<string, string> = kit().EQUATION_GALLERY_STRINGS[locale];
	const say = (key: string) => names[key] ?? translate(locale, key as LocalizationKey);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover dve-equation-gallery';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', anchor.getAttribute('aria-label') ?? '');
	pop.dataset.equationGallery = kind;
	pop.append(equationStyle());
	if (kind === 'builtIn')
		builtIns(pop, say, handlers.choose, anchor.dataset.equationGallery === 'insert-builtIn');
	else if (kind === 'symbols') symbols(pop, say, handlers.choose);
	else structures(pop, kind, say, handlers.choose);
	roving(pop);
	pop.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') handlers.dismissed?.();
	});
	if (!mountPopover(anchor, pop, anchor)) return;
	const first = pop.querySelector<HTMLElement>('[role="menuitem"]');
	if (first) first.tabIndex = 0;
	(pop.querySelector<HTMLElement>('select') ?? first)?.focus({ preventScroll: true });
}
