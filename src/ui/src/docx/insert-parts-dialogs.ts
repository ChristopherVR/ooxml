import { glossaryBlocksWritable, type Block } from 'ooxml-core/docx';
import { createOfficeGuid } from 'ooxml-core/crypto';
import { checkbox, dialogButton, fieldset, labelled, selectOf, textInput } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { formDialog, note, type InsertPartsHost } from './insert-parts-form';
import { translate, type LocalizationKey } from './localization';

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
	// Word saves new Quick Parts to Building Blocks.dotx; the document is the other choice.
	const saveIn = selectOf([
		['shared', 'Building Blocks'],
		['document', 'This document'],
	]);
	const saveInRow = labelled('Save in', saveIn);
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
				saveInRow,
			),
			message,
		],
		ok,
		cancel,
		first: name,
		inputs: [name, gallery, category, description, saveIn],
		reset(locale) {
			const shared = host.sharedGlossaryAvailable();
			saveInRow.hidden = !shared;
			saveIn.value = shared ? 'shared' : 'document';
			// A shared document does not take glossary changes; the shared store still does.
			saveIn.options[1]!.disabled = host.glossaryLocked();
			if (!shared && host.glossaryLocked()) {
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
			const entry = {
				name: name.value.trim(),
				gallery: gallery.value,
				category: category.value.trim(),
				...(description.value.trim() ? { description: description.value.trim() } : {}),
				blocks,
			};
			if (saveIn.value === 'shared' && host.sharedGlossaryAvailable()) host.saveSharedEntry(entry);
			else host.saveGlossaryEntry(entry);
			return true;
		},
	});
}

/** The first words of the selection, as the suggested name of a new building block. */
function firstWords(blocks: Block[]): string {
	const first = blocks.find((block) => block.type === 'paragraph');
	const text = first?.type === 'paragraph' ? first.runs.map((run) => run.text).join('') : '';
	return text.trim().replace(/\s+/g, ' ').slice(0, 32).trim();
}
