import {
	DOCUMENT_PROPERTY_CHOICES,
	glossaryEntryKey,
	type DocumentPropertyId,
	type GlossaryEntry,
	type HeaderFooterDesign,
} from 'ooxml-core/docx';
import { emit } from './events';
import { findLocalizedControl, localeOf, localizeElement } from './localization';
import type { RibbonAction } from './ribbon-action';
import { menuSelect, stack } from './ribbon-parts';
import { refitRibbon } from './ribbon-overflow';
import { panelOf } from './ribbon-tab-api';

/** The new commands borrow an existing icon's frame, then draw their own outline. */
const outline = (control: HTMLElement, d: string): void =>
	control.querySelector('svg path')?.setAttribute('d', d);

/** Header and footer designs, listed after Blank and before the first-page and even-page variants. */
const DESIGN_OPTIONS: Array<[string, string]> = [
	['threeColumns', 'Blank (Three Columns)'],
	['title', 'Title'],
	['titleAndAuthor', 'Title and Author'],
	['centeredTitle', 'Centered Title'],
	['authorAndPage', 'Author and Page Number'],
	['dateAndPage', 'Date and Page Number'],
];

function designAction(kind: 'header' | 'footer', value: string): RibbonAction {
	if (value === 'edit') return { type: 'headerFooterEdit', kind };
	if (value === 'remove') return { type: 'headerFooterRemove', kind };
	return { type: 'headerFooter', kind, slot: 'default', design: value as HeaderFooterDesign };
}

const QUICK_PARTS_MENU: Array<[string, string]> = [
	...DOCUMENT_PROPERTY_CHOICES.map((choice): [string, string] => [
		`property:${choice.id}`,
		`Document Property: ${choice.label}`,
	]),
	['field', 'Field…'],
	['save', 'Save Selection to Quick Part Gallery…'],
	['organizer', 'Building Blocks Organizer…'],
];

const QUICK_PARTS_DIALOGS = {
	field: 'fieldDialog',
	organizer: 'buildingBlocks',
	save: 'saveQuickPart',
} as const;

export function quickPartsMenuAction(value: string): RibbonAction {
	const [kind = '', rest = ''] = value.split(':');
	if (kind in QUICK_PARTS_DIALOGS)
		return {
			type: 'insertParts',
			kind: QUICK_PARTS_DIALOGS[kind as keyof typeof QUICK_PARTS_DIALOGS],
		};
	return {
		type: 'quickPart',
		command:
			kind === 'entry' || kind === 'shared'
				? {
						kind: 'entry',
						key: decodeURIComponent(rest),
						...(kind === 'shared' ? { shared: true } : {}),
					}
				: { kind: 'property', id: rest as DocumentPropertyId },
	};
}

/**
 * Insert > Signature Line: Signature Setup inserts the line; Sign signs the first unsigned line
 * (or the document) with the host's `documentSigner`, and is disabled without one.
 */
function signatureLineMenu(): HTMLElement {
	const wrap = menuSelect(
		'Signature line',
		'bookmark',
		[
			['setup', 'Signature Line…'],
			['sign', 'Sign…'],
		],
		(value) => ({ type: 'insertParts', kind: value === 'sign' ? 'sign' : 'signatureSetup' }),
		{ inline: true, momentary: true },
	);
	const sign = wrap.querySelector<HTMLOptionElement>('option[value="sign"]');
	if (sign) {
		sign.disabled = true;
		sign.dataset.signLine = '';
	}
	outline(wrap, 'M5 4h14v16H5z M8 10l5 5 M13 10l-5 5 M8 18h8');
	return wrap;
}

const groupLabelled = (insert: HTMLElement, label: string) =>
	[...insert.querySelectorAll<HTMLElement>('.ribbon-group')].find(
		(el) => el.dataset.label === label,
	);

/** The strings of languages other than English, loaded when one is shown. */
const translations = (locale: string): Promise<unknown> =>
	locale === 'en' ? Promise.resolve() : import('./locales/insert-parts-translations');

/**
 * Adds Quick Parts and Signature Line to the Insert tab's Text group, wires the Media group's
 * Online Video command, and adds the gallery designs to the Header and Footer menus. Called once,
 * after the ribbon is built; calling it again does nothing.
 */
export async function mountInsertParts(root: HTMLElement): Promise<void> {
	const insert = panelOf(root, 'insert');
	if (!insert || insert.dataset.insertParts !== undefined) return;
	insert.dataset.insertParts = '';
	const locale = localeOf(root);
	await translations(locale);
	const added: HTMLElement[] = [];
	const text = groupLabelled(insert, 'Text');
	if (text) {
		const quick = menuSelect('Quick Parts', 'coverPage', QUICK_PARTS_MENU, quickPartsMenuAction, {
			inline: true,
			momentary: true,
		});
		outline(quick, 'M5 4h14v16H5z M8 8h8 M8 12h8 M8 16h4 M16 15l2 2 M14 14l4 4');
		// Word's layout: two columns of three small commands, then Text Box and Content Control.
		const part = (label: string): HTMLElement | undefined => {
			const control = findLocalizedControl<HTMLElement>(text, label);
			return control?.closest<HTMLElement>('.ribbon-menu') ?? control ?? undefined;
		};
		const pick = (...items: Array<HTMLElement | undefined>) =>
			items.filter((item): item is HTMLElement => Boolean(item));
		// The commands of this group keep only their icons (names stay as tooltips and accessible names) so the group stays open at 1600.
		const iconOnly = (label: string): HTMLElement | undefined => {
			const item = part(label);
			if (!item) return undefined;
			item.classList.remove('ribbon-menu-inline');
			if (item.classList.contains('ribbon-menu')) item.classList.add('ribbon-menu-compact');
			item.querySelector(':scope > span')?.remove();
			item.title ||= label;
			item.style.width = '32px';
			item.style.minWidth = '0';
			return item;
		};
		// Word's two columns of Quick Parts, WordArt, Drop Cap | Signature Line, Date & Time, Object.
		const keep = (el: HTMLElement) => {
			el.classList.remove('ribbon-menu-inline');
			el.classList.add('ribbon-menu-compact');
			el.querySelector(':scope > span')?.remove();
			el.style.width = '32px';
			return el;
		};
		const sign = signatureLineMenu();
		added.push(keep(quick), keep(sign));
		const first = stack(...pick(quick, iconOnly('WordArt'), iconOnly('Drop cap')));
		const second = stack(...pick(sign, iconOnly('Date and time'), iconOnly('Object')));
		const third = stack(...pick(iconOnly('Text box'), iconOnly('Content control')));
		for (const old of text.querySelectorAll<HTMLElement>('.ribbon-stack'))
			if (!old.childElementCount) old.remove();
		text.prepend(first, second, third);
		added.push(first, second, third);
	}
	// The Media group is the ribbon's own; its command works once the host allows a video site.
	const media = groupLabelled(insert, 'Media');
	const video = media?.querySelector<HTMLButtonElement>('button');
	if (media && video) {
		video.dataset.action = JSON.stringify({ type: 'insertParts', kind: 'onlineVideo' });
		video.disabled = false;
		video.removeAttribute('data-unavailable');
		video.removeAttribute('data-localeariadescription');
		video.title = 'Online Video';
		media.dataset.onlineVideo = '';
		added.push(media);
	}
	for (const [kind, label] of [
		['header', 'Header'],
		['footer', 'Footer'],
	] as const) {
		const select = findLocalizedControl<HTMLSelectElement>(root, label);
		if (!select) continue;
		const anchor = select.options[1] ?? null;
		for (const [value, name] of DESIGN_OPTIONS)
			select.insertBefore(new Option(name, value), anchor);
		for (const [value, name] of [
			['edit', kind === 'header' ? 'Edit Header' : 'Edit Footer'],
			['remove', kind === 'header' ? 'Remove Header' : 'Remove Footer'],
		] as const)
			select.append(new Option(name, value));
		// A design is its own action; the menu's built-in choices keep their handler.
		select.parentElement?.addEventListener(
			'change',
			(event) => {
				if (event.target !== select || ['blank', 'first', 'even'].includes(select.value)) return;
				event.stopImmediatePropagation();
				emit(select, 'ribbon-action', designAction(kind, select.value));
				select.selectedIndex = -1;
			},
			true,
		);
		added.push(select);
	}
	for (const el of added) localizeElement(el, locale);
	// A language chosen later translates the whole ribbon, these controls included.
	new MutationObserver(() => {
		const next = localeOf(root);
		void translations(next).then(() => localizeElement(root, next));
	}).observe(root, { attributeFilter: ['data-editor-locale'] });
	syncInsertParts(root, undefined, []);
	root.dispatchEvent(new CustomEvent('insert-parts-ready'));
}

const MENU_ENTRIES = 12;

/**
 * Shows Online Video only when the host has named video sites to embed, and lists the document's
 * Quick Parts gallery entries at the top of the Quick Parts menu.
 */
export function syncInsertParts(
	toolbar: HTMLElement,
	glossary: readonly GlossaryEntry[] | undefined,
	videoHosts: readonly string[],
	shared: readonly GlossaryEntry[] = [],
	canSign = false,
): void {
	for (const option of toolbar.querySelectorAll<HTMLOptionElement>('option[data-sign-line]'))
		option.disabled = !canSign;
	// `hidden` alone loses to the group's own display rule.
	let changed = false;
	for (const el of toolbar.querySelectorAll<HTMLElement>('[data-online-video]')) {
		const hide = videoHosts.length === 0;
		changed ||= el.hidden !== hide;
		el.hidden = hide;
		el.style.display = hide ? 'none' : '';
	}
	// The ribbon decides which groups fold from the groups it shows.
	if (changed) refitRibbon(toolbar);
	const select = findLocalizedControl<HTMLSelectElement>(toolbar, 'Quick Parts');
	if (!select) return;
	// The document's entries first, then the shared ones (Word lists both in one gallery).
	const entries = [
		...(glossary ?? []).map((entry) => ({ entry, source: 'entry' })),
		...shared.map((entry) => ({ entry, source: 'shared' })),
	]
		.filter(({ entry }) => entry.gallery === 'docParts')
		.slice(0, MENU_ENTRIES);
	const signature = JSON.stringify(
		entries.map(({ entry, source }) => [source, glossaryEntryKey(entry), entry.description]),
	);
	if ((select.dataset.entries ?? '[]') === signature) return;
	select.dataset.entries = signature;
	for (const old of select.querySelectorAll('option[data-entry]')) old.remove();
	const first = select.options[0] ?? null;
	for (const { entry, source } of entries) {
		const option = new Option(
			entry.name,
			`${source}:${encodeURIComponent(glossaryEntryKey(entry))}`,
		);
		option.dataset.entry = '';
		if (entry.description) option.title = entry.description;
		select.insertBefore(option, first);
	}
	select.selectedIndex = -1;
}
