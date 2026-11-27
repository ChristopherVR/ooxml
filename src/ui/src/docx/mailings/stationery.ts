import {
	loadMailMergeStationery,
	type EnvelopeOptions,
	type LabelOptions,
} from 'ooxml-core/docx/mail-merge';
import { saveDocx, type DocumentModel } from 'ooxml-core/docx';
import { Selection } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	textInput,
} from '../dialog-fields';
import type { EditorCore } from '../editor-core';
import { openMergedDocument } from '../file-commands';
import type { EditorLocale } from '../localization';
import { modelToDoc } from '../model-adapter';
import { mailingsReason, type MailingsCommand } from '../mailings-ribbon';
import { dialogMessage, dialogSubmit, type MailingsDialog } from './dialog';

type StationeryEngine = Awaited<ReturnType<typeof loadMailMergeStationery>>;

export async function withStationery(
	core: EditorCore,
	key: MailingsCommand,
	use: (engine: StationeryEngine) => void,
): Promise<void> {
	const { view, loadGeneration } = core;
	try {
		const engine = await loadMailMergeStationery();
		if (core.view === view && core.loadGeneration === loadGeneration && !mailingsReason(core, key))
			use(engine);
	} catch (cause) {
		if (core.view === view && core.loadGeneration === loadGeneration) core.host.reportError(cause);
	}
}

export function emptyBody(core: EditorCore): boolean {
	const [block] = core.model.blocks;
	return (
		core.model.blocks.length === 1 &&
		block?.type === 'paragraph' &&
		!block.contentControl &&
		!block.bookmarks?.length &&
		!block.permissions?.length &&
		block.runs.length === 1 &&
		!block.runs[0]?.text &&
		!core.view?.state.doc.firstChild?.content.size &&
		!core.model.sections?.some((section) => section.headers || section.footers)
	);
}

function textarea(): HTMLTextAreaElement {
	const input = document.createElement('textarea');
	input.rows = 3;
	input.maxLength = 4096;
	return input;
}

export function envelopeDialog(
	dialog: MailingsDialog,
	locale: EditorLocale,
	envelopeSizes: StationeryEngine['envelopeSizes'],
	apply: (options: EnvelopeOptions) => void,
): void {
	const { t } = dialog;
	const delivery = textarea();
	const returnAddress = textarea();
	const omit = checkbox(t('Omit return address'));
	const size = selectOf(envelopeSizes.map(({ id, label }) => [id, label]));
	const offset = () => numberInput(0, 12, 0.01);
	const deliveryLeft = offset();
	const deliveryTop = offset();
	const returnLeft = offset();
	const returnTop = offset();
	const tray = numberInput(0, 65535, 1);
	const position = (left: HTMLInputElement, top: HTMLInputElement) => ({
		leftTwips: left.valueAsNumber * 1440,
		topTwips: top.valueAsNumber * 1440,
	});
	const showReturn = () => {
		returnAddress.disabled = omit.input.checked;
		returnOffsets.disabled = omit.input.checked;
	};
	const returnOffsets = fieldset(
		t('Return address position'),
		row(labelled(t('From left (inches)'), returnLeft), labelled(t('From top (inches)'), returnTop)),
	);
	omit.input.addEventListener('change', showReturn);
	dialog.content.append(
		labelled(t('Delivery address'), delivery),
		labelled(t('Return address'), returnAddress),
		omit.wrapper,
		labelled(t('Envelope size'), size),
		fieldset(
			t('Delivery address position'),
			row(
				labelled(t('From left (inches)'), deliveryLeft),
				labelled(t('From top (inches)'), deliveryTop),
			),
		),
		returnOffsets,
		labelled(t('First-page printer tray (optional)'), tray),
	);
	const controls = [deliveryLeft, deliveryTop, returnLeft, returnTop, tray];
	const ok = dialogSubmit(dialog, locale, () => {
		const selected = envelopeSizes.find(({ id }) => id === size.value);
		if (!selected || controls.some((input) => !input.checkValidity())) return;
		const { valueAsNumber } = tray;
		dialog.close();
		apply({
			size: selected.id,
			delivery: delivery.value,
			returnAddress: returnAddress.value,
			omitReturnAddress: omit.input.checked,
			deliveryPosition: position(deliveryLeft, deliveryTop),
			returnPosition: position(returnLeft, returnTop),
			...(Number.isFinite(valueAsNumber) && { firstPageTray: valueAsNumber }),
		});
	});
	for (const input of controls)
		input.addEventListener(
			'input',
			() => (ok.disabled = controls.some((item) => !item.checkValidity())),
		);
}

type LabelSpec = StationeryEngine['labelCatalogue'][number];

/** "2.63 in × 1 in" on a Letter sheet, "63.5 mm × 38.1 mm" on a metric one. */
export function labelSize(spec: LabelSpec): string {
	const inches = spec.pageWidth === 12_240;
	const value = (twips: number) =>
		inches
			? `${Math.round((twips / 1440) * 100) / 100} in`
			: `${Math.round((twips / 1440) * 254) / 10} mm`;
	return `${value(spec.width)} × ${value(spec.height)}`;
}

export function labelsDialog(
	dialog: MailingsDialog,
	engine: Pick<StationeryEngine, 'labelVendors' | 'searchLabels'>,
	merge: boolean,
	apply: (options: LabelOptions) => void,
): void {
	const { t } = dialog;
	const vendor = selectOf(engine.labelVendors.map((value) => [value, value]));
	const search = textInput();
	search.maxLength = 32;
	search.type = 'search';
	const product = selectOf([]);
	product.size = 6;
	const text = textarea();
	const mode = selectOf([
		['page', t('Full page of the same label')],
		['single', t('Single label')],
	]);
	const sheetRow = numberInput(1, 10, 1);
	const column = numberInput(1, 3, 1);
	const location = row(labelled(t('Label row'), sheetRow), labelled(t('Label column'), column));
	let listed: LabelSpec[] = [];
	const selected = () =>
		listed.find((option) => option.product === product.value && option.vendor === vendor.value);
	const output = dialogButton(t(merge ? 'Create merge template' : 'New Document'), true);
	const valid = () =>
		Boolean(selected()) &&
		(merge ||
			location.hidden ||
			(sheetRow.checkValidity() &&
				column.checkValidity() &&
				sheetRow.value !== '' &&
				column.value !== ''));
	const position = () => {
		const spec = selected();
		if (spec) {
			sheetRow.max = String(spec.rows);
			column.max = String(spec.columns);
			sheetRow.value = column.value = '1';
		}
		output.disabled = !valid();
	};
	const products = () => {
		listed = engine.searchLabels(search.value, vendor.value);
		product.replaceChildren(
			...listed.map(
				(spec) =>
					new Option(
						`${spec.product} (${spec.columns} × ${spec.rows}, ${labelSize(spec)})`,
						spec.product,
					),
			),
		);
		if (!listed.length) {
			const none = new Option(t('No matching products'), '');
			none.disabled = true;
			product.append(none);
		} else product.selectedIndex = 0;
		position();
	};
	vendor.addEventListener('change', products);
	search.addEventListener('input', products);
	product.addEventListener('change', position);
	dialog.content.append(
		labelled(t('Label vendor'), vendor),
		labelled(t('Find a product number'), search),
		labelled(t('Product number'), product),
	);
	if (!merge) {
		location.hidden = true;
		dialog.content.append(
			labelled(t('Label text'), text),
			labelled(t('Label output'), mode),
			location,
		);
	}
	products();
	mode.addEventListener('change', () => {
		location.hidden = mode.value !== 'single';
		output.disabled = !valid();
	});
	for (const input of [sheetRow, column])
		input.addEventListener('input', () => (output.disabled = !valid()));
	output.addEventListener('click', () => {
		const spec = selected();
		if (!spec || !valid()) return;
		dialog.close();
		apply({
			product: spec.product,
			vendor: spec.vendor,
			...(!merge && { text: text.value }),
			...(!merge &&
				mode.value === 'single' && {
					single: { row: sheetRow.valueAsNumber, column: column.valueAsNumber },
				}),
		});
	});
	dialog.actions(output);
}

export function applyDocument(
	core: EditorCore,
	model: DocumentModel,
	selectDelivery = false,
): void {
	const view = core.view;
	if (!view?.editable) return;
	const next = modelToDoc(model);
	const { doc } = view.state;
	const transaction = closeHistory(view.state.tr);
	if (!next.content.eq(doc.content)) transaction.replaceWith(0, doc.content.size, next.content);
	for (const key of [
		'pageWidth',
		'pageHeight',
		'marginTop',
		'marginRight',
		'marginBottom',
		'marginLeft',
		'sections',
		'sectionParts',
		'mailMerge',
	])
		if (next.attrs[key] !== doc.attrs[key]) transaction.setDocAttribute(key, next.attrs[key]);
	if (selectDelivery) {
		const position = next.lastChild ? next.content.size - next.lastChild.nodeSize : 0;
		transaction.setSelection(Selection.near(transaction.doc.resolve(position), 1));
	} else if (model.blocks[0]?.type === 'table')
		transaction.setSelection(Selection.atStart(transaction.doc));
	view.dispatch(transaction);
}

export function startLabels(core: EditorCore, { blocks, page, sections }: DocumentModel): void {
	applyDocument(core, {
		...core.model,
		blocks,
		page,
		...(sections && { sections }),
		mailMerge: { ...core.model.mailMerge, mainDocumentType: 'mailingLabels' },
	});
}

export async function openLabels(
	core: EditorCore,
	model: DocumentModel,
	current: () => boolean,
): Promise<void> {
	const bytes = await saveDocx(model);
	if (current()) openMergedDocument(core.element, bytes, 'labels.docx');
}

export function labelUpdateIssue(dialog: MailingsDialog): void {
	dialogMessage(dialog, 'Choose a simple label grid with a small first label.');
	dialog.actions();
}
