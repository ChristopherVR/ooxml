/**
 * Pure helpers of `office-ui-select` (choice data, popup placement, keyboard stepping). Moved
 * from pptx-viewer (`packages/shared/src/web-components/select-menu.ts`). The markup is the
 * element's own `render()`.
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
