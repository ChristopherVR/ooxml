import { wordDesignDefaults, type DesignEdit, type WordDesignScheme } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import { applyDesignEdit } from './design-commands';
import { customTitles, openDesignDialog } from './design-dialog';
import { translate, type LocalizationKey } from './localization';
import { canEditDesign, type DesignMenu } from './design-ribbon-controls';
import { changeDesignScheme, readDesignSchemes } from './design-scheme-store';
import { browseTheme, saveTheme } from './design-theme-files';
import { customEntry, designEntries, type Entry, type Presets } from './design-gallery-entries';
import styles from './design-gallery.css?raw';

function saveDefault(core: EditorCore): void {
	if (!core.wordDesignPreferences || !canEditDesign(core)) return;
	const defaults = wordDesignDefaults(core.model);
	try {
		if (!defaults) throw new Error(translate(core.locale, 'design.defaultError'));
		core.wordDesignPreferences.write(defaults);
	} catch {
		core.host.reportError(translate(core.locale, 'design.defaultError'));
	}
}

function appendItem(menu: HTMLElement, label: string, choose: () => void, entry?: Entry) {
	const item = document.createElement('button');
	item.type = 'button';
	item.setAttribute('aria-label', label);
	item.setAttribute('role', entry ? 'menuitemradio' : 'menuitem');
	if (entry) {
		item.setAttribute('aria-checked', String(entry.active));
		if (entry.unavailable) {
			item.setAttribute('aria-disabled', 'true');
			item.setAttribute('aria-description', entry.unavailable);
			item.title = entry.unavailable;
		}
		entry.preview.setAttribute('aria-hidden', 'true');
		item.append(entry.preview);
	} else item.className = 'gallery-command';
	const text = document.createElement('span');
	text.className = 'gallery-name';
	text.textContent = label;
	item.append(text);
	if (!entry?.unavailable) item.addEventListener('click', choose);
	menu.append(item);
	return item;
}

function galleryGroup(menu: HTMLElement, name: string): HTMLElement {
	const group = document.createElement('div');
	group.setAttribute('role', 'group');
	group.setAttribute('aria-label', name);
	const title = document.createElement('div');
	title.className = 'dve-design-group-label';
	title.textContent = name;
	title.setAttribute('aria-hidden', 'true');
	group.append(title);
	menu.append(group);
	return group;
}

export function openDocumentDesign(
	core: EditorCore,
	anchor: HTMLButtonElement,
	kind: DesignMenu,
	presets: Presets,
): void {
	if (!canEditDesign(core)) return;
	if (kind === 'default') {
		openDesignDialog(core, anchor, kind, () => saveDefault(core));
		return;
	}
	const menu = document.createElement('div');
	menu.className = 'ribbon-popover ribbon-gallery-menu dve-design-menu';
	menu.setAttribute('role', 'menu');
	menu.setAttribute('aria-label', anchor.getAttribute('aria-label') ?? '');
	const style = document.createElement('style');
	style.textContent = styles;
	menu.append(style);
	const pick = (apply: () => void) => {
		closeRibbonPopover();
		anchor.focus();
		apply();
	};
	const say = (key: LocalizationKey) => translate(core.locale, key);
	const items: HTMLButtonElement[] = [];
	const pickEntry = (group: HTMLElement, entry: Entry) =>
		items.push(
			appendItem(
				group,
				entry.label,
				() => {
					const { edit } = entry;
					if (edit) pick(() => applyDesignEdit(core, edit));
				},
				entry,
			),
		);
	if (kind === 'colors' || kind === 'fonts') {
		const saved = readDesignSchemes(core).filter((scheme) => scheme.kind === kind);
		if (saved.length) {
			const group = galleryGroup(menu, say('Custom'));
			for (const scheme of saved) pickEntry(group, customEntry(core, scheme));
			for (const scheme of saved) {
				items.push(
					appendItem(
						group,
						say('design.editScheme').replace('{name}', () => scheme.name),
						() =>
							pick(() =>
								openDesignDialog(core, anchor, kind, (edit) => saveScheme(edit, scheme), {
									scheme,
									remove: () => changeDesignScheme(core, scheme, undefined),
								}),
							),
					),
				);
			}
		}
	}
	function saveScheme(edit?: DesignEdit, previous?: WordDesignScheme): boolean {
		if (!edit) return false;
		if (
			(edit.kind === 'colors' || edit.kind === 'fonts') &&
			core.wordDesignPreferences?.schemes &&
			!changeDesignScheme(core, previous, edit)
		)
			return false;
		applyDesignEdit(core, edit);
		return true;
	}
	const builtIn =
		kind === 'theme' || kind === 'colors' || kind === 'fonts'
			? galleryGroup(menu, kind === 'theme' ? 'Office' : say('design.builtIn'))
			: menu;
	if (kind === 'theme') builtIn.classList.add('dve-design-theme-grid');
	for (const entry of designEntries(core, kind, presets)) pickEntry(builtIn, entry);
	if (kind === 'theme') {
		items.push(
			appendItem(menu, say('design.browseThemes'), () => pick(() => browseTheme(core, anchor))),
			appendItem(menu, say('design.saveTheme'), () => pick(() => saveTheme(core, anchor))),
		);
	}
	if (kind === 'colors' || kind === 'fonts' || kind === 'spacing') {
		items.push(
			appendItem(menu, translate(core.locale, customTitles[kind]), () =>
				pick(() => openDesignDialog(core, anchor, kind, saveScheme)),
			),
		);
	}
	menu.addEventListener('keydown', (event) => {
		const at = items.findIndex((item) => item === event.target);
		const next = {
			ArrowDown: at + 1,
			ArrowUp: at - 1,
			ArrowRight: at + 1,
			ArrowLeft: at - 1,
			Home: 0,
			End: items.length - 1,
		}[event.key];
		if (next === undefined) return;
		event.preventDefault();
		items[(next + items.length) % items.length]?.focus();
	});
	if (mountPopover(anchor, menu, anchor)) {
		const active = items.find((item) => item.getAttribute('aria-checked') === 'true');
		(active ?? items[0])?.focus();
	}
}
