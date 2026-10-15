/** File > Options > Customize Ribbon: show or hide every ribbon command. */
import type { RibbonControl } from 'ooxml-core/xlsx/ui';
import { isMenuSeparator, ribbonTabs } from 'ooxml-core/xlsx/ui';
import type { BackstageHost } from './parts';

/** Every command id a control (and its menu) runs. */
function commandsOf(control: RibbonControl): string[] {
	switch (control.kind) {
		case 'stack':
			return control.controls.flatMap(commandsOf);
		case 'separator':
			return [];
		case 'menu':
			return control.items.flatMap((item) => (isMenuSeparator(item) ? [] : [item.command]));
		case 'split':
			return [
				control.command,
				...control.menu.flatMap((item) => (isMenuSeparator(item) ? [] : [item.command])),
			];
		default:
			return [control.command];
	}
}

/**
 * Customize Ribbon, reduced to showing and hiding: every tab, group and command with a checkbox.
 * Unchecked commands go into `hiddenActions` (and the `ribbon-customize` event). Drawn into
 * `pane`, the slot the Options dialog shows for the Customize Ribbon category.
 */
export function renderCustomizeRibbon(pane: HTMLElement, host: BackstageHost): void {
	const { t, commands } = host.ctx;
	const doc = pane.ownerDocument;
	const hidden = new Set(host.hiddenActions());
	const boxes: Array<{ id: string; input: HTMLInputElement }> = [];
	const commit = () =>
		host.setHiddenActions(boxes.filter((box) => !box.input.checked).map((box) => box.id));
	const seen = new Set<string>();
	const sections: HTMLElement[] = [];
	for (const tab of ribbonTabs()) {
		const section = doc.createElement('section');
		const title = doc.createElement('h3');
		title.textContent = t(tab.label);
		section.append(title);
		for (const group of tab.groups) {
			const fieldset = doc.createElement('fieldset');
			fieldset.className = 'xve-customize-group';
			const legend = doc.createElement('legend');
			legend.textContent = t(group.label);
			fieldset.append(legend);
			for (const id of group.controls.flatMap(commandsOf)) {
				const command = commands.get(id);
				if (!command || seen.has(id)) continue;
				seen.add(id);
				const row = doc.createElement('label');
				row.className = 'xve-customize-row';
				const input = doc.createElement('input');
				input.type = 'checkbox';
				input.checked = !hidden.has(id);
				input.addEventListener('change', commit);
				const text = doc.createElement('span');
				text.textContent = t(command.label);
				row.append(input, text);
				fieldset.append(row);
				boxes.push({ id, input });
			}
			if (fieldset.childElementCount > 1) section.append(fieldset);
		}
		if (section.childElementCount > 1) sections.push(section);
	}
	const note = doc.createElement('p');
	note.className = 'xve-options-note';
	note.textContent = t('Choose which commands the ribbon shows.');
	const reset = doc.createElement('button');
	reset.type = 'button';
	reset.className = 'xve-options-action';
	reset.textContent = t('Reset all customizations');
	reset.addEventListener('click', () => {
		for (const box of boxes) box.input.checked = true;
		commit();
	});
	pane.replaceChildren(note, reset, ...sections);
}
