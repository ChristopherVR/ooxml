import {
	glossaryBlocksWritable,
	glossaryEntryKey,
	glossaryPreview,
	type Block,
	type GlossaryEntry,
} from 'ooxml-core/docx';
import { createOfficeGuid } from 'ooxml-core/crypto';
import { checkbox, dialogButton, fieldset, labelled, selectOf, textInput } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { formDialog, note, type InsertPartsHost } from './insert-parts-form';
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

export type { InsertPartsHost } from './insert-parts-form';

/** Insert > Signature Line > Signature Setup. */
export function createSignatureSetupDialog(host: InsertPartsHost): FormatDialog {
	const signer = textInput();
	const title = textInput();
	const email = textInput();
	const instructions = document.createElement('textarea');
	instructions.rows = 3;
	const comments = checkbox('Allow the signer to add comments in the Sign dialog');
	const date = checkbox('Show sign date in signature line');
	for (const input of [signer, title, email, instructions]) input.maxLength = 500;
	const ok = dialogButton('OK', true);
	const cancel = dialogButton('Cancel');
	return formDialog(host, {
		title: 'Signature Setup',
		className: 'dve-signature-dialog',
		content: [
			labelled('Suggested signer', signer),
			labelled("Suggested signer's title", title),
			labelled("Suggested signer's e-mail address", email),
			labelled('Instructions to the signer', instructions),
			comments.wrapper,
			date.wrapper,
			note(
				'This editor inserts the signature line. Signing it is not available here: it needs a certificate and a signing service.',
			),
		],
		ok,
		cancel,
		first: signer,
		inputs: [signer, title, email, instructions],
		reset() {
			for (const input of [signer, title, email, instructions]) input.value = '';
			comments.input.checked = false;
			date.input.checked = true;
			return true;
		},
		valid: () => true,
		apply() {
			const optional = (key: string, value: string) =>
				value.trim() ? { [key]: value.trim() } : {};
			host.insertSignatureLine({
				id: createOfficeGuid(),
				...optional('suggestedSigner', signer.value),
				...optional('suggestedSignerTitle', title.value),
				...optional('suggestedSignerEmail', email.value),
				...optional('instructions', instructions.value),
				allowComments: comments.input.checked,
				showSignDate: date.input.checked,
			});
			return true;
		},
	});
}

/** Insert > Quick Parts > Save Selection to Quick Part Gallery (and AutoText). */
export function createSaveQuickPartDialog(host: InsertPartsHost): FormatDialog {
	const name = textInput();
	const gallery = selectOf([
		['docParts', 'Quick Parts'],
		['autoTxt', 'AutoText'],
	]);
	const category = textInput();
	const description = textInput();
	for (const input of [name, category, description]) input.maxLength = 255;
	const message = note('');
	let blocks: Block[] | undefined;
	const ok = dialogButton('OK', true);
	const cancel = dialogButton('Cancel');
	return formDialog(host, {
		title: 'Create New Building Block',
		className: 'dve-save-quick-part-dialog',
		content: [
			fieldset(
				'Building block',
				labelled('Name', name),
				labelled('Gallery', gallery),
				labelled('Category', category),
				labelled('Description', description),
			),
			message,
		],
		ok,
		cancel,
		first: name,
		inputs: [name, gallery, category, description],
		reset(locale) {
			if (host.glossaryLocked()) {
				host.report(
					translate(
						locale,
						'Building blocks cannot be changed while the document is shared.' as LocalizationKey,
					),
				);
				return false;
			}
			blocks = host.selectionBlocks();
			if (!blocks) return false;
			name.value = firstWords(blocks);
			category.value = 'General';
			description.value = '';
			message.textContent = glossaryBlocksWritable(blocks)
				? ''
				: translate(
						locale,
						'This selection cannot be saved to a gallery yet: pictures, links and notes need parts the gallery does not have.' as LocalizationKey,
					);
			return true;
		},
		valid: () =>
			Boolean(
				blocks && glossaryBlocksWritable(blocks) && name.value.trim() && category.value.trim(),
			),
		apply() {
			if (!blocks) return false;
			host.saveGlossaryEntry({
				name: name.value.trim(),
				gallery: gallery.value,
				category: category.value.trim(),
				...(description.value.trim() ? { description: description.value.trim() } : {}),
				blocks,
			});
			return true;
		},
	});
}

const GALLERY_NAMES = { docParts: 'Quick Parts', autoTxt: 'AutoText' } as const;

/** Insert > Quick Parts > Building Blocks Organizer, for the document's own glossary. */
export function createBuildingBlocksDialog(host: InsertPartsHost): FormatDialog {
	const list = document.createElement('select');
	list.size = 10;
	list.style.minHeight = '13em';
	const preview = document.createElement('output');
	preview.className = 'dve-building-block-preview';
	const insert = dialogButton('Insert', true);
	const remove = dialogButton('Delete');
	const close = dialogButton('Close');
	let entries: GlossaryEntry[] = [];
	let locale: EditorLocale = 'en';
	const fill = (selectKey?: string) => {
		// Entries Word keeps for its own use (placeholders of content controls) are not listed.
		entries = host.glossary().filter((entry) => entry.gallery !== 'placeholder');
		list.replaceChildren(
			...entries.map((entry) => {
				const gallery =
					entry.gallery in GALLERY_NAMES
						? translate(
								locale,
								GALLERY_NAMES[entry.gallery as keyof typeof GALLERY_NAMES] as LocalizationKey,
							)
						: entry.gallery;
				return new Option(`${entry.name} (${gallery}, ${entry.category})`, glossaryEntryKey(entry));
			}),
		);
		if (selectKey) list.value = selectKey;
		else if (entries.length) list.selectedIndex = 0;
		sync();
	};
	const selected = () => entries.find((entry) => glossaryEntryKey(entry) === list.value);
	const sync = () => {
		const entry = selected();
		preview.textContent = entry
			? [glossaryPreview(entry), entry.description].filter(Boolean).join(' - ')
			: entries.length
				? ''
				: 'This document has no building blocks.';
		insert.disabled = !entry;
		remove.disabled = !entry || host.glossaryLocked();
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
		const entry = selected();
		if (entry && host.editable()) {
			hide();
			host.insertEntry(entry);
		}
	});
	remove.addEventListener('click', () => {
		const entry = selected();
		if (!entry || !host.editable()) return;
		host.removeGlossaryEntry(glossaryEntryKey(entry));
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
		},
	};
}

/** The first words of the selection, as the suggested name of a new building block. */
function firstWords(blocks: Block[]): string {
	const first = blocks.find((block) => block.type === 'paragraph');
	const text = first?.type === 'paragraph' ? first.runs.map((run) => run.text).join('') : '';
	return text.trim().replace(/\s+/g, ' ').slice(0, 32).trim();
}
