/** File > Options (General, Formulas) and Customize Ribbon. */
import type { RibbonControl } from '../ribbon/parts';
import { isMenuSeparator, ribbonTabs } from '../ribbon/parts';
import { heading, labelled, paragraph, primary, type PageContext } from './parts';

const LANGUAGES: Array<[string, string]> = [
	['en', 'English'],
	['fr', 'Français'],
	['de', 'Deutsch'],
	['es', 'Español'],
	['zh-CN', '简体中文'],
];

export function renderOptions(page: PageContext): void {
	const { t, host, doc } = page;
	const current = host.options();
	const language = doc.createElement('select');
	for (const [value, label] of LANGUAGES) language.append(new Option(label, value));
	language.value = current.locale;
	language.addEventListener('change', () => host.setOption('locale', language.value));
	const theme = doc.createElement('select');
	for (const [value, label] of [
		['auto', 'Automatic'],
		['light', 'Light'],
		['dark', 'Dark'],
	] as const)
		theme.append(new Option(t(label), value));
	theme.value = current.theme === 'light' || current.theme === 'dark' ? current.theme : 'auto';
	theme.addEventListener('change', () => host.setOption('theme', theme.value));
	const author = doc.createElement('input');
	author.type = 'text';
	author.value = current.author;
	author.addEventListener('change', () => host.setOption('author', author.value));
	const calculation = doc.createElement('select');
	calculation.append(new Option(t('Automatic'), 'automatic'), new Option(t('Manual'), 'manual'));
	calculation.value = current.calculation;
	calculation.addEventListener('change', () =>
		host.setOption('calculation', calculation.value === 'manual' ? 'manual' : 'automatic'),
	);
	const r1c1 = doc.createElement('label');
	r1c1.className = 'xve-backstage-check';
	const box = doc.createElement('input');
	box.type = 'checkbox';
	box.disabled = true;
	const text = doc.createElement('span');
	text.textContent = t('R1C1 reference style (not available)');
	r1c1.append(box, text);
	page.content.replaceChildren(
		heading(page, t('Options')),
		heading(page, t('General'), 'h3'),
		labelled(page, t('Display language'), language),
		labelled(page, t('Theme'), theme),
		labelled(page, t('Author name'), author),
		heading(page, t('Formulas'), 'h3'),
		labelled(page, t('Workbook calculation'), calculation),
		paragraph(
			page,
			t('Manual: formulas recalculate when you press F9 (Calculate Now).'),
			'xve-backstage-muted',
		),
		r1c1,
		paragraph(
			page,
			t('Options apply to this editor and are not stored between sessions.'),
			'xve-backstage-muted',
		),
	);
}

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
 * Unchecked commands go into `hiddenActions` (and the `ribbon-customize` event).
 */
export function renderCustomize(page: PageContext): void {
	const { t, host, doc } = page;
	const { commands } = host.ctx;
	const hidden = new Set(host.hiddenActions());
	const boxes: Array<{ id: string; input: HTMLInputElement }> = [];
	const commit = () =>
		host.setHiddenActions(boxes.filter((box) => !box.input.checked).map((box) => box.id));
	const seen = new Set<string>();
	const sections: HTMLElement[] = [];
	for (const tab of ribbonTabs()) {
		const section = doc.createElement('section');
		section.append(heading(page, t(tab.label), 'h3'));
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
	const reset = primary(page, t('Reset all customizations'), () => {
		for (const box of boxes) box.input.checked = true;
		commit();
	});
	page.content.replaceChildren(
		heading(page, t('Customize Ribbon')),
		paragraph(page, t('Choose which commands the ribbon shows.'), 'xve-backstage-muted'),
		reset,
		...sections,
	);
}
