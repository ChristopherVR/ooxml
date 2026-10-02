/** File > New, Open, Save As, Print and Export. */
import { saveExtension, withExtension } from '../file-commands';
import { printRange } from '../print';
import {
	card,
	facts,
	heading,
	labelled,
	paragraph,
	primary,
	runFile,
	type PageContext,
} from './parts';
import { TEMPLATES } from './templates';
import { formatRange } from '@christophervr/xlsx-core';

function tile(
	page: PageContext,
	label: string,
	run: () => void,
	template?: string,
): HTMLButtonElement {
	const { doc } = page;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = 'xve-backstage-tile';
	const preview = doc.createElement('span');
	preview.className = 'xve-backstage-tile-preview';
	if (template) preview.dataset.template = template;
	const text = doc.createElement('span');
	text.textContent = label;
	button.append(preview, text);
	button.setAttribute('aria-label', label);
	button.addEventListener('click', run);
	return button;
}

export function renderNew(page: PageContext): void {
	const { t, host, doc } = page;
	const tiles = doc.createElement('div');
	tiles.className = 'xve-backstage-tiles';
	tiles.append(tile(page, t('Blank workbook'), runFile(page, 'new')));
	for (const template of TEMPLATES) {
		const button = tile(
			page,
			t(template.label),
			() => {
				host.close();
				host.newFromTemplate(template.id);
			},
			template.id,
		);
		button.title = t(template.description);
		tiles.append(button);
	}
	page.content.replaceChildren(heading(page, t('New')), tiles);
}

export function renderOpen(page: PageContext): void {
	const { t } = page;
	page.content.replaceChildren(
		heading(page, t('Open')),
		paragraph(
			page,
			t('Open an Excel workbook (.xlsx, .xlsm, .xls) or a CSV file from this device.'),
		),
		primary(page, t('Browse...'), runFile(page, 'open')),
		paragraph(page, t('Files are processed entirely in the browser.'), 'xve-backstage-muted'),
	);
}

export function renderSaveAs(page: PageContext): void {
	const { t, host, doc } = page;
	const name = doc.createElement('input');
	name.type = 'text';
	name.value = host.fileName().replace(/\.[^.]+$/, '');
	const format = doc.createElement('select');
	const workbookExtension = saveExtension(host.fileName());
	format.append(
		new Option(
			t(
				workbookExtension === 'xlsm'
					? 'Excel Macro-Enabled Workbook (*.xlsm)'
					: 'Excel Workbook (*.xlsx)',
			),
			workbookExtension,
		),
		new Option(t('CSV UTF-8 (Comma delimited) (*.csv), current sheet only'), 'csv'),
	);
	const save = () => {
		const base = name.value.trim();
		if (!base) return;
		const csv = format.value === 'csv';
		runFile(
			page,
			csv ? 'exportCsv' : 'saveAs',
			withExtension(base, csv ? 'csv' : workbookExtension),
		)();
	};
	name.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') save();
	});
	page.content.replaceChildren(
		heading(page, t('Save As')),
		paragraph(page, t('Save a copy of this workbook to your device.'), 'xve-backstage-muted'),
		labelled(page, t('File name'), name),
		labelled(page, t('Format'), format),
		primary(page, t('Save'), save),
	);
}

export function renderPrint(page: PageContext): void {
	const { t, host } = page;
	const workbook = host.ctx.workbook();
	const sheet = host.ctx.activeSheet();
	const area = workbook ? printRange(workbook, sheet) : undefined;
	const hasPrintArea = Boolean(workbook?.sheets[sheet]?.pageSetup?.printArea);
	page.content.replaceChildren(
		heading(page, t('Print')),
		primary(page, t('Print'), runFile(page, 'print')),
		facts(page, [
			['Sheet', workbook?.sheets[sheet]?.name ?? '-'],
			[hasPrintArea ? 'Print area' : 'Used range', area ? formatRange(area) : t('Empty sheet')],
		]),
		paragraph(
			page,
			t(
				'Printing uses the browser print dialog and prints the active sheet as a table. Page breaks, headers, footers and scaling are approximate.',
			),
			'xve-backstage-muted',
		),
	);
}

export function renderExport(page: PageContext): void {
	const { t } = page;
	page.content.replaceChildren(
		heading(page, t('Export')),
		card(
			page,
			t('Create a PDF'),
			t('Opens the print dialog; choose Save as PDF as the destination.'),
			primary(page, t('Create PDF'), runFile(page, 'print')),
		),
		card(
			page,
			t('Save a copy as XLSX'),
			t('Download the workbook as an Excel .xlsx file.'),
			primary(page, t('Save a copy as XLSX'), runFile(page, 'export')),
		),
		card(
			page,
			t('Export the current sheet as CSV'),
			t('Download the values of the current sheet as CSV UTF-8, as the cells display them.'),
			primary(page, t('Export as CSV'), runFile(page, 'exportCsv')),
		),
	);
}
