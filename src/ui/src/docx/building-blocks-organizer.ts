import { glossaryEntryKey, glossaryPreview, type GlossaryEntry } from 'ooxml-core/docx';
import { dialogButton, labelled } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import type { InsertPartsHost } from './insert-parts-form';
import {
	localizeElement,
	translate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';

const GALLERY_NAMES = { docParts: 'Quick Parts', autoTxt: 'AutoText' } as const;

/** Where an Organizer entry lives: the document's glossary or the shared building blocks. */
type Source = 'document' | 'shared';
interface Listed {
	entry: GlossaryEntry;
	source: Source;
}

const valueOf = ({ entry, source }: Listed) => `${source}:${glossaryEntryKey(entry)}`;

/**
 * Insert > Quick Parts > Building Blocks Organizer: the document's own entries and, when the host
 * keeps them, the building blocks shared by every document (Word's Building Blocks.dotx), each
 * labelled with where it is stored.
 */
export function createBuildingBlocksDialog(host: InsertPartsHost): FormatDialog {
	const list = document.createElement('select');
	list.size = 10;
	list.style.minHeight = '13em';
	const preview = document.createElement('output');
	preview.className = 'dve-building-block-preview';
	const insert = dialogButton('Insert', true);
	const remove = dialogButton('Delete');
	const close = dialogButton('Close');
	let entries: Listed[] = [];
	let locale: EditorLocale = 'en';
	const t = (key: string) => translate(locale, key as LocalizationKey);
	const fill = (selectValue?: string) => {
		// Entries Word keeps for its own use (placeholders of content controls) are not listed.
		const visible = (entry: GlossaryEntry) => entry.gallery !== 'placeholder';
		entries = [
			...host
				.glossary()
				.filter(visible)
				.map((entry): Listed => ({ entry, source: 'document' })),
			...host
				.sharedGlossary()
				.filter(visible)
				.map((entry): Listed => ({ entry, source: 'shared' })),
		];
		list.replaceChildren(
			...entries.map((item) => {
				const { entry, source } = item;
				const gallery =
					entry.gallery in GALLERY_NAMES
						? t(GALLERY_NAMES[entry.gallery as keyof typeof GALLERY_NAMES])
						: entry.gallery;
				const where = t(source === 'shared' ? 'Building Blocks' : 'This document');
				return new Option(`${entry.name} (${gallery}, ${entry.category}, ${where})`, valueOf(item));
			}),
		);
		if (selectValue && entries.some((item) => valueOf(item) === selectValue))
			list.value = selectValue;
		else if (entries.length) list.selectedIndex = 0;
		sync();
	};
	const selected = () => entries.find((item) => valueOf(item) === list.value);
	const sync = () => {
		const item = selected();
		preview.textContent = item
			? [glossaryPreview(item.entry), item.entry.description].filter(Boolean).join(' - ')
			: entries.length
				? ''
				: host.sharedGlossaryAvailable()
					? 'There are no building blocks.'
					: 'This document has no building blocks.';
		insert.disabled = !item;
		remove.disabled = !item || (item.source === 'document' && host.glossaryLocked());
		localizeElement(preview, locale);
	};
	list.addEventListener('change', sync);
	list.addEventListener('dblclick', () => insert.click());
	const element = createDialogShell('Building Blocks Organizer', 'dve-format-dialog dve-organizer');
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		host.focus();
	};
	insert.addEventListener('click', () => {
		const item = selected();
		if (item && host.editable()) {
			hide();
			host.insertEntry(item.entry);
		}
	});
	remove.addEventListener('click', () => {
		const item = selected();
		if (!item || !host.editable()) return;
		const key = glossaryEntryKey(item.entry);
		if (item.source === 'shared') host.removeSharedEntry(key);
		else host.removeGlossaryEntry(key);
		fill();
	});
	close.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	const actions = dialogActions(remove, close, insert);
	return {
		element,
		close: hide,
		get isOpen() {
			return isDialogOpen(element);
		},
		setLocale(next) {
			locale = next;
		},
		open() {
			element.replaceChildren(labelled('Building blocks', list), preview, actions);
			fill();
			localizeDialog(element, locale);
			setDialogOpen(element, true);
			list.focus();
			// Another document may have changed the shared entries since they were read.
			if (host.sharedGlossaryAvailable())
				void host.loadSharedGlossary().then(() => {
					if (isDialogOpen(element)) fill(list.value);
				});
		},
	};
}
