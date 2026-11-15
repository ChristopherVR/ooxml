import {
	inspectDocument,
	INSPECTION_CATEGORIES,
	type CategoryInspection,
	type DocumentModel,
	type InspectionCategory,
	type InspectionFindingKind,
} from 'ooxml-core/docx';
import { checkbox, dialogButton } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';

/** What the Document Inspector needs from the editor. */
export interface DocumentInspectorHost {
	model(): DocumentModel;
	/** Why Remove All is unavailable for `category`, or undefined when it can run. */
	blocked(category: InspectionCategory): 'readOnly' | 'protected' | undefined;
	/** Removes the category as one undoable step; false when nothing was removed. */
	remove(category: InspectionCategory): boolean;
	restoreFocus(): void;
}

/** Word's category names and descriptions, in English (translated by the dialog's localizer). */
export const CATEGORY_TEXT: Record<InspectionCategory, readonly [string, string]> = {
	comments: [
		'Comments, Revisions, and Versions',
		'Inspects the document for comments and revision marks.',
	],
	properties: [
		'Document Properties and Personal Information',
		'Inspects for hidden metadata or personal information saved with the document.',
	],
	customXml: ['Custom XML Data', 'Inspects for custom XML data stored with this document.'],
	headersFooters: [
		'Headers, Footers, and Watermarks',
		'Inspects the document for information in headers, footers, and watermarks.',
	],
	hiddenText: ['Hidden Text', 'Inspects the document for text that has been formatted as hidden.'],
};

export const FINDING_TEXT: Record<InspectionFindingKind, string> = {
	comments: 'Comments',
	revisions: 'Revision marks',
	documentProperties: 'Document properties',
	author: 'Author',
	reviewers: 'Reviewer names',
	company: 'Company',
	manager: 'Manager',
	hyperlinkBase: 'Hyperlink base',
	template: 'Template name',
	customProperties: 'Custom properties',
	customXml: 'Custom XML data',
	headers: 'Headers',
	footers: 'Footers',
	watermarks: 'Watermarks',
	hiddenText: 'Hidden text',
};

const element = <K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text = '') => {
	const node = document.createElement(tag);
	if (className) node.className = className;
	if (text) node.textContent = text;
	return node;
};

/** The bold category name, with an id the row's Remove All button is described by. */
function categoryText(category: InspectionCategory): HTMLElement {
	const text = element('div', 'dve-inspector-text');
	const title = element('strong', 'dve-inspector-title', CATEGORY_TEXT[category][0]);
	title.id = `dve-inspector-${category}`;
	text.append(title);
	return text;
}

/** File > Info > Check for Issues > Inspect Document: choose categories, inspect, Remove All. */
export function createDocumentInspectorDialog(host: DocumentInspectorHost): FormatDialog {
	const dialog = createDialogShell('Document Inspector', 'dve-format-dialog dve-inspector-dialog');
	let locale: EditorLocale = 'en';
	const selected = new Set<InspectionCategory>(INSPECTION_CATEGORIES);
	/** Categories removed since the last inspection, which then read as cleared. */
	const removed = new Set<InspectionCategory>();
	let results: CategoryInspection[] | undefined;

	const inspectButton = dialogButton('Inspect', true);
	const reinspect = dialogButton('Reinspect');
	reinspect.classList.add('dve-inspector-reinspect');
	const closeButton = dialogButton('Close');

	const close = () => {
		setDialogOpen(dialog, false);
		dialog.replaceChildren();
		results = undefined;
		host.restoreFocus();
	};

	const choices = () => {
		const list = element('div', 'dve-inspector-list');
		for (const category of INSPECTION_CATEGORIES) {
			const { wrapper, input } = checkbox(CATEGORY_TEXT[category][0]);
			wrapper.classList.add('dve-inspector-row');
			wrapper.querySelector('span')?.remove();
			input.checked = selected.has(category);
			input.dataset.category = category;
			input.addEventListener('change', () => {
				if (input.checked) selected.add(category);
				else selected.delete(category);
				inspectButton.disabled = selected.size === 0;
			});
			const text = categoryText(category);
			text.append(element('span', 'dve-inspector-description', CATEGORY_TEXT[category][1]));
			wrapper.append(text);
			list.append(wrapper);
		}
		return list;
	};

	const result = (entry: CategoryInspection) => {
		const row = element('div', 'dve-inspector-row dve-inspector-result');
		row.dataset.category = entry.category;
		const cleared = removed.has(entry.category);
		const found = entry.findings.length > 0 && !cleared;
		row.dataset.state = found ? 'found' : 'clear';
		const mark = element('span', 'dve-inspector-mark', found ? '!' : '✓');
		mark.setAttribute('aria-hidden', 'true');
		const text = categoryText(entry.category);
		row.append(mark, text);
		if (!found) {
			const message = cleared ? 'All items were removed.' : 'No items were found.';
			text.append(element('span', 'dve-inspector-description', message));
			return row;
		}
		text.append(element('span', 'dve-inspector-description', 'The following items were found:'));
		const items = element('ul', 'dve-inspector-findings');
		for (const finding of entry.findings)
			items.append(element('li', '', FINDING_TEXT[finding.kind]));
		text.append(items);
		const blocked = host.blocked(entry.category);
		if (blocked === 'protected')
			text.append(
				element(
					'span',
					'dve-inspector-note',
					'Removing these items is not allowed while the document is protected.',
				),
			);
		const button = dialogButton('Remove All');
		button.disabled = Boolean(blocked);
		button.setAttribute('aria-describedby', `dve-inspector-${entry.category}`);
		button.addEventListener('click', () => {
			if (!host.remove(entry.category)) return;
			removed.add(entry.category);
			render();
			reinspect.focus();
		});
		row.append(button);
		return row;
	};

	function render() {
		let body: HTMLElement;
		if (results) {
			body = element('div', 'dve-inspector-list');
			for (const entry of results) body.append(result(entry));
		} else body = choices();
		const intro = results
			? 'Review the inspection results.'
			: 'To check the document for the selected content, click Inspect.';
		dialog.replaceChildren(
			element('p', 'dve-inspector-intro', intro),
			body,
			dialogActions(results ? reinspect : inspectButton, closeButton),
		);
		inspectButton.disabled = selected.size === 0;
		localizeDialog(dialog, locale);
	}

	inspectButton.addEventListener('click', () => {
		removed.clear();
		results = inspectDocument(host.model(), [...selected]);
		render();
		reinspect.focus();
	});
	reinspect.addEventListener('click', () => {
		results = undefined;
		render();
		inspectButton.focus();
	});
	closeButton.addEventListener('click', close);
	onDialogDismiss(dialog, close);
	return {
		element: dialog,
		close,
		get isOpen() {
			return isDialogOpen(dialog);
		},
		setLocale(next) {
			locale = next;
			if (isDialogOpen(dialog)) render();
		},
		open() {
			results = undefined;
			removed.clear();
			render();
			setDialogOpen(dialog, true);
			inspectButton.focus();
		},
	};
}
