// The Equation tab's galleries: the built-in equations, the symbols by category and the eleven
// structure galleries, each drawn as the equation it inserts (empty arguments as placeholders).
import { EQUATION_GALLERY } from 'ooxml-core/math';
import { equationKit as kit } from './equation-kit';
import { equationPreview, renderEquation } from '../math/equation-render';
import type { RibbonAction } from './ribbon-action';
import { equationStyle } from './equation-editing';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import { localeOf, translate, type LocalizationKey } from './localization';

type Choose = (action: RibbonAction) => void;

function item(pop: HTMLElement, label: string, choose: () => void): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.setAttribute('role', 'menuitem');
	button.setAttribute('aria-label', label);
	button.title = label;
	// The equation editor keeps its caret while a gallery item is pressed.
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', () => {
		closeRibbonPopover();
		choose();
	});
	pop.append(button);
	return button;
}

function preview(button: HTMLElement, latex: string, label: string): void {
	const omml = kit().buildUpEquation(latex, 'latex', false);
	const drawn = omml ? equationPreview(omml, true) : undefined;
	const math = drawn && renderEquation(button.ownerDocument, drawn, false, label);
	if (math) button.append(math);
	else button.textContent = label;
}

/** Arrow keys move between a gallery's items, as in Word's galleries. */
function roving(pop: HTMLElement): void {
	pop.addEventListener('keydown', (event) => {
		const items = [...pop.querySelectorAll<HTMLElement>('[role="menuitem"]')];
		const at = items.indexOf(event.target as HTMLElement);
		if (at < 0) return;
		const step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[event.key];
		if (!step) return;
		event.preventDefault();
		items[(at + step + items.length) % items.length]?.focus();
	});
}

function structures(pop: HTMLElement, group: string, say: (key: string) => string, choose: Choose) {
	const gallery = kit().EQUATION_STRUCTURES.find((entry) => entry.id === group);
	for (const structure of gallery?.structures ?? []) {
		const label = say(structure.name);
		preview(
			item(pop, label, () =>
				choose({ type: 'equation', key: 'structure', value: `${group}/${structure.id}` }),
			),
			structure.latex,
			label,
		);
	}
}

function builtIns(pop: HTMLElement, say: (key: string) => string, choose: Choose) {
	for (const entry of EQUATION_GALLERY) {
		const label = say(entry.name);
		const button = item(pop, label, () =>
			choose({ type: 'equation', key: 'builtIn', value: entry.id }),
		);
		button.classList.add('dve-equation-gallery-wide');
		const name = document.createElement('span');
		name.className = 'gallery-name';
		name.textContent = label;
		button.append(name);
		preview(button, entry.latex, label);
	}
	const add = item(pop, say('Insert New Equation'), () => choose({ type: 'equation', key: 'new' }));
	add.classList.add('dve-equation-gallery-command');
	add.textContent = say('Insert New Equation');
}

function symbols(pop: HTMLElement, say: (key: string) => string, choose: Choose) {
	const category = document.createElement('select');
	category.setAttribute('aria-label', say('Symbol category'));
	category.className = 'dve-equation-symbol-category';
	for (const entry of kit().EQUATION_SYMBOLS) category.add(new Option(say(entry.name), entry.id));
	const grid = document.createElement('div');
	grid.className = 'dve-equation-symbol-grid';
	grid.setAttribute('role', 'group');
	const show = () => {
		grid.replaceChildren();
		const entry =
			kit().EQUATION_SYMBOLS.find((c) => c.id === category.value) ?? kit().EQUATION_SYMBOLS[0]!;
		grid.setAttribute('aria-label', say(entry.name));
		for (const symbol of entry.symbols)
			item(grid, symbol, () => choose({ type: 'equation', key: 'symbol', value: symbol })).append(
				symbol,
			);
	};
	category.addEventListener('change', show);
	pop.append(category, grid);
	show();
}

/** Opens gallery `kind` (`builtIn`, `symbols` or a structure group) under `anchor`. */
export function openEquationGallery(anchor: HTMLElement, kind: string, choose: Choose): void {
	const locale = localeOf(anchor);
	// Structure and symbol category names come with the galleries; the rest is the editor's.
	const names: Record<string, string> = kit().EQUATION_GALLERY_STRINGS[locale];
	const say = (key: string) => names[key] ?? translate(locale, key as LocalizationKey);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover dve-equation-gallery';
	pop.setAttribute('role', 'menu');
	pop.setAttribute('aria-label', anchor.getAttribute('aria-label') ?? '');
	pop.dataset.equationGallery = kind;
	pop.append(equationStyle());
	if (kind === 'builtIn') builtIns(pop, say, choose);
	else if (kind === 'symbols') symbols(pop, say, choose);
	else structures(pop, kind, say, choose);
	roving(pop);
	if (mountPopover(anchor, pop, anchor))
		pop.querySelector<HTMLElement>('select, [role="menuitem"]')?.focus({ preventScroll: true });
}
