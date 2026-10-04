import { createIconSvg, paintIcon } from './icons.js';

/** One Quick Access Toolbar button; every string arrives translated from the host. */
export interface OfficeQuickAccessItem {
	id: string;
	/** Registered glyph name (see `registerIcon`). */
	icon: string;
	/** Accessible name, and the visible text when labels show. */
	label: string;
	/** ScreenTip; absent leaves the button without a `title`. */
	title?: string | undefined;
	disabled?: boolean | undefined;
}

export interface QuickAccessStrip {
	toolbar: HTMLElement;
	render(items: readonly OfficeQuickAccessItem[], label: string, showLabels: boolean): void;
}

/**
 * The Quick Access Toolbar: one roving tab stop, arrows/Home/End between enabled buttons.
 * Buttons are keyed by id and patched in place, so a focused button survives every update.
 */
export function createQuickAccessStrip(
	doc: Document,
	activate: (id: string) => void,
): QuickAccessStrip {
	const toolbar = doc.createElement('div');
	toolbar.className = 'qat';
	toolbar.setAttribute('role', 'toolbar');
	toolbar.setAttribute('part', 'quick-access');
	const buttons = new Map<string, HTMLButtonElement>();
	let stop = '';
	const enabled = () => [...toolbar.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
	const syncRoving = (): void => {
		const list = enabled();
		const current = list.find((b) => b.dataset.command === stop) ?? list[0];
		for (const button of buttons.values()) button.tabIndex = button === current ? 0 : -1;
	};
	toolbar.addEventListener('focusin', (event) => {
		const target = (event.target as Element).closest('button');
		if (target?.dataset.command) {
			stop = target.dataset.command;
			syncRoving();
		}
	});
	toolbar.addEventListener('keydown', (event) => {
		if (event.ctrlKey || event.metaKey || event.altKey) return;
		const list = enabled();
		const index = list.findIndex((b) => b === (event.target as Element).closest('button'));
		const last = list.length - 1;
		const moves: Record<string, number> = {
			ArrowRight: (index + 1) % list.length,
			ArrowLeft: (index + last) % list.length,
			Home: 0,
			End: last,
		};
		const next = moves[event.key] ?? -1;
		if (next >= 0 && index >= 0) {
			event.preventDefault();
			list[next]?.focus();
		}
		// Keep arrows and native activation out of the host document's shortcuts.
		if (next >= 0 || event.key === ' ' || event.key === 'Enter') event.stopPropagation();
	});

	const create = (id: string): HTMLButtonElement => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.dataset.command = id;
		button.append(createIconSvg(doc));
		button.addEventListener('click', () => activate(id));
		return button;
	};

	return {
		toolbar,
		render(items, label, showLabels) {
			toolbar.hidden = items.length === 0;
			toolbar.setAttribute('aria-label', label);
			for (const id of [...buttons.keys()]) {
				if (!items.some((item) => item.id === id)) {
					buttons.get(id)?.remove();
					buttons.delete(id);
				}
			}
			for (const item of items) {
				let button = buttons.get(item.id);
				if (!button) {
					button = create(item.id);
					buttons.set(item.id, button);
				}
				paintIcon(button.querySelector('svg')!, item.icon);
				button.setAttribute('aria-label', item.label);
				if (item.title === undefined) button.removeAttribute('title');
				else button.title = item.title;
				button.disabled = item.disabled === true;
				let text = button.querySelector('small');
				if (showLabels) {
					text ??= button.appendChild(doc.createElement('small'));
					text.textContent = item.label;
				} else {
					text?.remove();
				}
			}
			// Reorder only when the order really changed: moving a focused node blurs it.
			const current = [...toolbar.children];
			if (items.some((item, index) => current[index] !== buttons.get(item.id)))
				toolbar.append(...items.map((item) => buttons.get(item.id)!));
			syncRoving();
		},
	};
}
