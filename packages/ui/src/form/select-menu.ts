/**
 * Popup helpers of `office-ui-select`. Moved from pptx-viewer
 * (`packages/shared/src/web-components/select-menu.ts`, `select-trigger.ts`).
 */

/** Mutations that can change a select's value, labels or font preview. */
export const SELECT_OPTION_ATTRIBUTES = [
	'value',
	'disabled',
	'label',
	'selected',
	'hidden',
	'data-display-label',
	'data-description',
	'style',
];

/** The flattened option data the popup renders. */
export interface SelectChoice {
	value: string;
	label: string;
	disabled: boolean;
	hidden: boolean;
	group: string;
	displayLabel?: string;
	description?: string;
	fontFamily?: string;
}

/** Flatten `<option>` children, keeping optgroup labels and disabled or hidden state. */
export function collectSelectChoices(options: HTMLOptionElement[]): SelectChoice[] {
	return options.map((option) => {
		const parent = option.parentElement;
		const group = parent?.localName === 'optgroup' ? (parent as HTMLOptGroupElement) : null;
		return {
			value: option.value,
			label: option.label || option.textContent?.trim() || '',
			disabled: option.disabled || Boolean(group?.disabled),
			hidden: Boolean(option.hidden) || Boolean(group?.hidden),
			group: group?.label ?? '',
			...(option.dataset.displayLabel ? { displayLabel: option.dataset.displayLabel } : {}),
			...(option.dataset.description ? { description: option.dataset.description } : {}),
			...(option.style.fontFamily ? { fontFamily: option.style.fontFamily } : {}),
		};
	});
}

/** Build the popup's options (only while it is open or has been opened). */
export function renderSelectMenu(menu: HTMLElement, choices: SelectChoice[], value: string): void {
	const doc = menu.ownerDocument;
	const items: HTMLElement[] = [];
	let previousGroup = '';
	choices.forEach((choice, index) => {
		if (choice.hidden) return;
		if (choice.group && choice.group !== previousGroup) {
			const heading = doc.createElement('div');
			heading.className = 'group';
			heading.setAttribute('role', 'presentation');
			heading.textContent = choice.group;
			items.push(heading);
		}
		previousGroup = choice.group;
		const item = doc.createElement('div');
		item.className = 'option';
		item.setAttribute('role', 'option');
		item.setAttribute('aria-selected', String(choice.value === value));
		item.setAttribute('aria-disabled', String(choice.disabled));
		item.id = `${menu.id}-${index}`;
		item.dataset.index = String(index);
		item.textContent = choice.displayLabel ?? choice.label;
		if (choice.fontFamily) item.style.fontFamily = choice.fontFamily;
		if (choice.description) {
			const description = doc.createElement('span');
			description.className = 'description';
			description.textContent = choice.description;
			item.append(description);
		}
		items.push(item);
	});
	menu.replaceChildren(...items);
}

/** A length token resolved on `el`, in pixels (falls back when unset or not in px). */
function tokenPx(el: Element, name: string, fallback: number): number {
	const value = el.ownerDocument.defaultView?.getComputedStyle(el).getPropertyValue(name).trim();
	const px = value?.endsWith('px') ? parseFloat(value) : NaN;
	return Number.isFinite(px) ? px : fallback;
}

/**
 * Prefer below the trigger, flipping above when that offers more room. Margins, the gap and the
 * height cap come from `--office-space-2`, `--office-space-1` and `--office-select-max-height`
 * (`--office-font-picker-max-height` for a font family picker).
 */
export function positionSelectMenu(
	menu: HTMLElement,
	trigger: HTMLElement,
	host: HTMLElement,
): void {
	const view = menu.ownerDocument.defaultView!;
	const rect = trigger.getBoundingClientRect();
	const margin = tokenPx(host, '--office-space-2', 8);
	const gap = tokenPx(host, '--office-space-1', 4);
	const width = Math.max(0, view.innerWidth - margin * 2);
	const below = Math.max(0, view.innerHeight - rect.bottom - gap - margin);
	const above = Math.max(0, rect.top - gap - margin);
	menu.style.minWidth = `${Math.min(rect.width, width)}px`;
	menu.style.maxWidth = `${width}px`;
	const cap =
		host.getAttribute('data-font-picker') === 'family'
			? tokenPx(host, '--office-font-picker-max-height', 320)
			: tokenPx(host, '--office-select-max-height', 240);
	// Measure at the normal cap before choosing a side; reset it so a constrained menu can grow.
	menu.style.maxHeight = `${cap}px`;
	const preferred = menu.getBoundingClientRect().height;
	const flip = below < preferred && above > below;
	menu.style.maxHeight = `${Math.min(cap, flip ? above : below)}px`;
	const { height, width: menuWidth } = menu.getBoundingClientRect();
	menu.style.top = `${Math.max(margin, flip ? rect.top - gap - height : rect.bottom + gap)}px`;
	menu.style.left = `${Math.max(margin, Math.min(rect.left, view.innerWidth - menuWidth - margin))}px`;
}

/** Expose the keyboard target while keeping the selected state separate. */
export function markSelectActive(
	menu: HTMLElement,
	trigger: HTMLElement,
	active: number,
	open: boolean,
): void {
	for (const item of menu.querySelectorAll<HTMLElement>('[data-index]'))
		item.toggleAttribute('data-active', Number(item.dataset.index) === active && open);
	if (open && active >= 0) trigger.setAttribute('aria-activedescendant', `${menu.id}-${active}`);
	menu.querySelector<HTMLElement>('[data-active]')?.scrollIntoView?.({ block: 'nearest' });
}

/** The next enabled, visible option when moving with the arrow keys (wrapping). */
export function nextSelectActive(choices: SelectChoice[], active: number, step: number): number {
	for (let i = 0; i < choices.length; i++) {
		active = (active + step + choices.length) % choices.length;
		if (!choices[active]!.disabled && !choices[active]!.hidden) return active;
	}
	return active;
}

/** Options a PageUp/PageDown press moves, like the visible rows of a native listbox. */
export const SELECT_PAGE_SIZE = 8;

/** Jump a page of options up (-1) or down (+1), clamped to the nearest enabled option. */
export function pageSelectActive(
	choices: SelectChoice[],
	active: number,
	direction: 1 | -1,
): number {
	const usable = (index: number) => !choices[index]?.disabled && !choices[index]?.hidden;
	const last = choices.length - 1;
	const target = Math.max(0, Math.min(last, Math.max(active, 0) + direction * SELECT_PAGE_SIZE));
	for (let index = target; index >= 0 && index <= last; index -= direction)
		if (usable(index)) return index;
	return active;
}

/** Trigger with an `icon` slot, the value text and a chevron. */
export function createSelectTrigger(doc: Document): {
	trigger: HTMLButtonElement;
	text: HTMLSpanElement;
} {
	const trigger = doc.createElement('button');
	trigger.type = 'button';
	trigger.setAttribute('part', 'trigger');
	trigger.setAttribute('role', 'combobox');
	trigger.setAttribute('aria-haspopup', 'listbox');
	trigger.setAttribute('aria-expanded', 'false');
	const icon = doc.createElement('slot');
	icon.name = 'icon';
	const text = doc.createElement('span');
	text.className = 'value';
	text.setAttribute('part', 'value');
	const chevron = doc.createElement('span');
	chevron.className = 'chevron';
	chevron.setAttribute('part', 'indicator');
	chevron.setAttribute('aria-hidden', 'true');
	const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('fill', 'none');
	svg.setAttribute('stroke', 'currentColor');
	svg.setAttribute('stroke-width', '2');
	svg.setAttribute('stroke-linecap', 'round');
	svg.setAttribute('stroke-linejoin', 'round');
	const path = doc.createElementNS(svg.namespaceURI, 'path');
	path.setAttribute('d', 'm6 9 6 6 6-6');
	svg.append(path);
	chevron.append(svg);
	trigger.append(icon, text, chevron);
	return { trigger, text };
}

/** Optional editable content (`slot="custom"`) stays in the light DOM so hosts own its events. */
export function prependSelectCustomSlot(host: HTMLElement, menu: HTMLElement): void {
	if (!host.querySelector('[slot="custom"]')) return;
	const custom = host.ownerDocument.createElement('slot');
	custom.name = 'custom';
	menu.prepend(custom);
}
