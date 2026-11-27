/**
 * The Options step of File > Export > Create a PDF, as Word's "Publish as PDF" Options dialog:
 * page range, what to publish, the outline, properties and tags, and PDF/A. Built on the shared
 * modal (`office-ui-dialog`: focus kept inside and returned, Escape cancels). Word's "Optimize for"
 * has no counterpart: pictures always keep their own resolution.
 */
import { createDialogSession } from '../../dialog/dialog-session';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	selectOf,
} from '../dialog-fields';
import { createDialogShell, dialogActions, localizeDialog, setDialogOpen } from '../dialog-shell';
import type { EditorLocale } from '../localization';
import type { EditorPdfOptions } from './editor-pdf';
import type { DocxPdfMarkupBalloons } from 'ooxml-core/docx/pdf-export';
import {
	defaultPdfChoice,
	pdfOptionsFromChoice,
	type PdfOptionsChoice,
	type PdfOptionsContext,
	type PdfRangeChoice,
} from './pdf-options-form';

export interface PdfOptionsDialog {
	element: HTMLElement;
	/** Shows the step; resolves to the options to export with, or undefined when cancelled. */
	choose(context: PdfOptionsContext): Promise<EditorPdfOptions | undefined>;
	setLocale(locale: EditorLocale): void;
	readonly isOpen: boolean;
}

let groups = 0;

/** Word's Review > Show Markup > Balloons choices, for "Document showing markup". */
const BALLOONS: ReadonlyArray<readonly [DocxPdfMarkupBalloons, string]> = [
	['commentsAndFormatting', 'Show only comments and formatting in balloons'],
	['all', 'Show revisions in balloons'],
	['comments', 'Show all revisions inline'],
	['none', 'Show no balloons'],
];

function radio(name: string, value: string, label: string) {
	const wrapper = document.createElement('label');
	wrapper.className = 'dve-dialog-check';
	const input = document.createElement('input');
	input.type = 'radio';
	input.name = name;
	input.value = value;
	input.setAttribute('aria-label', label);
	const text = document.createElement('span');
	text.textContent = label;
	wrapper.append(input, text);
	return { wrapper, input };
}

export function createPdfOptionsDialog(): PdfOptionsDialog {
	const element = createDialogShell('PDF options', 'dve-format-dialog dve-pdf-options-dialog');
	const group = `dve-pdf-${++groups}`;
	const ranges = new Map<PdfRangeChoice, HTMLInputElement>();
	const rangeRow = (value: PdfRangeChoice, label: string) => {
		const choice = radio(`${group}-range`, value, label);
		ranges.set(value, choice.input);
		return choice.wrapper;
	};
	const from = numberInput(1, 9999, 1);
	const to = numberInput(1, 9999, 1);
	const pagesRow = document.createElement('div');
	pagesRow.className = 'dve-dialog-row dve-pdf-pages';
	pagesRow.append(rangeRow('pages', 'Pages'), labelled('From', from), labelled('To', to));
	const documentOnly = radio(`${group}-what`, 'document', 'Document');
	const withMarkup = radio(`${group}-what`, 'markup', 'Document showing markup');
	const balloons = selectOf(BALLOONS);
	const balloonsField = labelled('Balloons', balloons);
	balloonsField.className = 'dve-pdf-balloons';
	const bookmarks = checkbox('Create bookmarks using');
	const headings = radio(`${group}-bookmarks`, 'headings', 'Headings');
	const wordBookmarks = radio(`${group}-bookmarks`, 'wordBookmarks', 'Word bookmarks');
	const sources = document.createElement('div');
	sources.className = 'dve-pdf-bookmark-sources';
	sources.append(headings.wrapper, wordBookmarks.wrapper);
	const properties = checkbox('Document properties');
	const tags = checkbox('Document structure tags for accessibility');
	const pdfA = checkbox('PDF/A compliant');
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('Export', true);
	const content = [
		fieldset(
			'Page range',
			rangeRow('all', 'All'),
			rangeRow('current', 'Current page'),
			rangeRow('selection', 'Selection'),
			pagesRow,
		),
		fieldset('Publish what', documentOnly.wrapper, withMarkup.wrapper, balloonsField),
		fieldset(
			'Include non-printing information',
			bookmarks.wrapper,
			sources,
			properties.wrapper,
			tags.wrapper,
		),
		fieldset('PDF options', pdfA.wrapper),
		dialogActions(cancel, confirm),
	];
	let locale: EditorLocale = 'en';
	let open = false;
	// Word keeps the last choices for the session; the page range starts at All each time.
	let last = defaultPdfChoice();

	const read = (): PdfOptionsChoice => ({
		range: [...ranges].find(([, input]) => input.checked)?.[0] ?? 'all',
		from: from.value,
		to: to.value,
		markup: withMarkup.input.checked,
		balloons: balloons.value as DocxPdfMarkupBalloons,
		createBookmarks: bookmarks.input.checked,
		bookmarkSource: wordBookmarks.input.checked ? 'wordBookmarks' : 'headings',
		documentProperties: properties.input.checked,
		taggedPdf: tags.input.checked,
		pdfA: pdfA.input.checked,
	});
	const syncSources = () => {
		headings.input.disabled = wordBookmarks.input.disabled = !bookmarks.input.checked;
		balloons.disabled = !withMarkup.input.checked;
	};
	for (const input of [bookmarks.input, documentOnly.input, withMarkup.input])
		input.addEventListener('change', syncSources);
	for (const field of [from, to])
		field.addEventListener('input', () => {
			const pages = ranges.get('pages');
			if (pages) pages.checked = true;
		});

	const fill = (choice: PdfOptionsChoice, context: PdfOptionsContext) => {
		const current = ranges.get('current')!;
		const selection = ranges.get('selection')!;
		current.disabled = !context.pages;
		selection.disabled = !context.hasSelection;
		for (const [value, input] of ranges) input.checked = value === choice.range;
		const page = String(context.pages?.current ?? 1);
		from.value = to.value = page;
		if (context.pages) from.max = to.max = String(context.pages.total);
		else from.max = to.max = '9999';
		documentOnly.input.checked = !choice.markup;
		withMarkup.input.checked = choice.markup;
		balloons.value = choice.balloons;
		bookmarks.input.checked = choice.createBookmarks;
		headings.input.checked = choice.bookmarkSource === 'headings';
		wordBookmarks.input.checked = choice.bookmarkSource === 'wordBookmarks';
		properties.input.checked = choice.documentProperties;
		tags.input.checked = choice.taggedPdf;
		pdfA.input.checked = choice.pdfA;
		syncSources();
	};

	return {
		element,
		choose(context) {
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			fill({ ...last, range: 'all' }, context);
			const session = createDialogSession<EditorPdfOptions>({
				element,
				confirm,
				cancel,
				submit: () => {
					const choice = read();
					const options = pdfOptionsFromChoice(choice, context);
					if ('invalid' in options) {
						const field = options.invalid === 'from' ? from : to;
						field.setAttribute('aria-invalid', 'true');
						field.focus();
						field.select();
						return undefined;
					}
					from.removeAttribute('aria-invalid');
					to.removeAttribute('aria-invalid');
					last = choice;
					return options;
				},
				finish: () => {
					open = false;
					setDialogOpen(element, false);
					from.removeAttribute('aria-invalid');
					to.removeAttribute('aria-invalid');
				},
				enter: (event) => (event.target as Element | null)?.tagName === 'INPUT',
			});
			open = true;
			setDialogOpen(element, true);
			ranges.get('all')?.focus();
			return session.result;
		},
		setLocale(next) {
			locale = next;
			if (open) localizeDialog(element, locale);
		},
		get isOpen() {
			return open;
		},
	};
}
