import { panelsOf } from './ribbon-tab-api';
import { ribbonControlId } from './ribbon-visibility';
import type { BackstageHandlers } from './backstage-pages';

/**
 * Word's Customize Ribbon, reduced to showing and hiding: every tab, group and command with a
 * checkbox. Unchecked commands go into `hiddenActions`; groups and tabs left empty disappear.
 * Drawn into `host`, the pane the File > Options dialog shows for the Customize Ribbon category.
 */
export function renderCustomizeRibbon(
	host: HTMLElement,
	handlers: BackstageHandlers,
	t: (text: string) => string,
): void {
	const ribbon = handlers.ribbon();
	const hidden = new Set(handlers.hiddenActions());
	const seen = new Set<string>();
	const boxes: Array<{ id: string; input: HTMLInputElement }> = [];
	const commit = () =>
		handlers.setHiddenActions(boxes.filter((b) => !b.input.checked).map((b) => b.id));
	const sections: HTMLElement[] = [];
	for (const panel of ribbon ? panelsOf(ribbon) : []) {
		const tabSection = document.createElement('section');
		tabSection.className = 'dve-customize-tab';
		const title = document.createElement('h3');
		title.textContent = panel.dataset.label ?? '';
		tabSection.append(title);
		for (const group of panel.querySelectorAll<HTMLElement>('.ribbon-group')) {
			const list = document.createElement('fieldset');
			list.className = 'dve-customize-group';
			const legend = document.createElement('legend');
			legend.textContent = group.dataset.caption ?? group.dataset.label ?? '';
			list.append(legend);
			for (const control of group.querySelectorAll<HTMLElement>(
				'button[aria-label]:not(.ribbon-launcher):not(.ribbon-overflow-button):not(.style-tile), select[aria-label], input[aria-label]',
			)) {
				const id = ribbonControlId(control);
				if (!id || seen.has(id)) continue;
				seen.add(id);
				const row = document.createElement('label');
				row.className = 'dve-customize-row';
				const input = document.createElement('input');
				input.type = 'checkbox';
				input.checked = !hidden.has(id);
				input.addEventListener('change', commit);
				const text = document.createElement('span');
				text.textContent = control.getAttribute('aria-label') ?? id;
				row.append(input, text);
				list.append(row);
				boxes.push({ id, input });
			}
			if (list.children.length > 1) tabSection.append(list);
		}
		if (tabSection.children.length > 1) sections.push(tabSection);
	}
	const note = document.createElement('p');
	note.className = 'dve-options-note';
	note.textContent = t('Choose which commands the ribbon shows.');
	const reset = document.createElement('button');
	reset.type = 'button';
	reset.className = 'dve-options-action';
	reset.textContent = t('Reset all customizations');
	reset.addEventListener('click', () => {
		for (const box of boxes) box.input.checked = true;
		commit();
	});
	host.replaceChildren(note, reset, ...sections);
}
