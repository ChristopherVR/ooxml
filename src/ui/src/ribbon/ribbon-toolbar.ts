import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import css from './ribbon-toolbar.css?raw';

const ITEM =
	'office-ui-button, office-ui-menu-button, office-ui-select, office-ui-checkbox, office-ui-switch';

/**
 * `role="toolbar"` container. Left/Right/Home/End move focus between enabled direct or
 * group-nested controls (Up/Down instead when `aria-orientation="vertical"`).
 * All items stay in the tab order: roving tabindex is not implemented yet.
 */
export class OfficeUiToolbar extends OfficeElement {
	static override styles = controlStyles(css);

	constructor() {
		super();
		this.addEventListener('keydown', (event) => this.onKey(event));
	}

	private items(): HTMLElement[] {
		return [...this.querySelectorAll<HTMLElement>(ITEM)].filter(
			(el) => !el.hasAttribute('disabled'),
		);
	}

	private onKey(event: KeyboardEvent): void {
		const vertical = this.getAttribute('aria-orientation') === 'vertical';
		const next = vertical ? 'ArrowDown' : 'ArrowRight';
		const prev = vertical ? 'ArrowUp' : 'ArrowLeft';
		if (![next, prev, 'Home', 'End'].includes(event.key)) return;
		const items = this.items();
		const from = items.findIndex((el) => el === event.target || el.contains(event.target as Node));
		if (from < 0 || items.length === 0) return;
		event.preventDefault();
		const index =
			event.key === 'Home'
				? 0
				: event.key === 'End'
					? items.length - 1
					: (from + (event.key === next ? 1 : -1) + items.length) % items.length;
		items[index]?.focus();
	}

	protected override willUpdate(): void {
		this.setAttribute('role', 'toolbar');
	}

	protected override render() {
		return html`<slot></slot>`;
	}
}

export const defineToolbar = definer('office-ui-toolbar', () => OfficeUiToolbar);
