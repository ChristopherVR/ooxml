/** File > Info: name and save state, compatibility notes, properties and workbook facts. */
import type { Workbook, WorkbookProperties } from '@christophervr/xlsx-core';
import { facts, heading, labelled, noteList, paragraph, primary, type PageContext } from './parts';
import { customProperties } from './properties-custom';

type T = PageContext['t'];

/** Notes about what this editor does with the opened format (shown on Info and on load). */
export function formatNotes(workbook: Workbook, t: T): string[] {
	if (workbook.format === 'xls')
		return [
			t(
				'This workbook was opened from an Excel 97-2003 file (.xls). Save writes an Excel Workbook (.xlsx); the original .xls file is not changed.',
			),
		];
	if (workbook.format === 'csv')
		return [
			t(
				'This workbook was opened from a CSV file, which keeps values only. Save writes an Excel Workbook (.xlsx); use Export to write CSV.',
			),
		];
	if (workbook.format === 'xlsm')
		return [t('Macros are kept in the file but do not run in this editor.')];
	return [];
}

/** Every note for Info: the format notes, then what the loader reported. */
export function compatibilityNotes(workbook: Workbook, t: T): string[] {
	return [...formatNotes(workbook, t), ...workbook.warnings];
}

type TextProperty = {
	[K in keyof WorkbookProperties]-?: WorkbookProperties[K] extends string | undefined ? K : never;
}[keyof WorkbookProperties];
const PROPERTY_FIELDS: ReadonlyArray<readonly [TextProperty, string]> = [
	['title', 'Title'],
	['subject', 'Subject'],
	['creator', 'Author'],
	['keywords', 'Tags'],
	['description', 'Comments'],
	['category', 'Category'],
	['contentStatus', 'Content status'],
	['company', 'Company'],
	['manager', 'Manager'],
	['hyperlinkBase', 'Hyperlink base'],
];

const FORMAT_NAMES: Record<Workbook['format'], string> = {
	xlsx: 'Excel Workbook (.xlsx)',
	xlsm: 'Excel Macro-Enabled Workbook (.xlsm)',
	xls: 'Excel 97-2003 Workbook (.xls)',
	csv: 'CSV (Comma delimited)',
	new: 'New workbook',
};

function countCells(workbook: Workbook): { cells: number; formulas: number } {
	let cells = 0;
	let formulas = 0;
	for (const sheet of workbook.sheets)
		for (const row of sheet.rows.values())
			for (const cell of row.values()) {
				cells++;
				if (cell.formula) formulas++;
			}
	return { cells, formulas };
}

export function renderInfo(page: PageContext): void {
	const { host, t, doc, content } = page;
	const workbook = host.ctx.workbook();
	const state = host.saveState();
	const status =
		state === 'dirty'
			? 'Unsaved changes'
			: state === 'saved-local'
				? 'Saved to this device'
				: 'Saved';
	const title = heading(page, host.fileName());
	title.classList.add('xve-backstage-title');
	const left = doc.createElement('div');
	left.append(
		title,
		paragraph(page, t(status), 'xve-backstage-muted'),
		heading(page, t('Inspect Workbook'), 'h3'),
		paragraph(page, t('Compatibility notes'), 'xve-backstage-muted'),
		noteList(
			page,
			workbook ? compatibilityNotes(workbook, t) : [],
			t('No compatibility notes for this workbook.'),
		),
		paragraph(
			page,
			t(
				'Charts are shown but not edited; pivot tables, sparklines, macros and other unsupported parts are kept in the file but not shown.',
			),
			'xve-backstage-muted',
		),
	);
	const right = doc.createElement('aside');
	right.className = 'xve-backstage-properties';
	right.append(heading(page, t('Properties'), 'h3'));
	if (workbook) {
		if (host.setPassword) {
			const protectedSave = host.passwordProtected?.() ?? false;
			const encrypt = primary(
				page,
				t(protectedSave ? 'Change password' : 'Encrypt Workbook'),
				() => void host.setPassword?.(),
			);
			encrypt.disabled = host.ctx.readOnly();
			left.append(
				heading(page, t('Protect Workbook'), 'h3'),
				paragraph(
					page,
					t(
						protectedSave
							? 'Excel saves are encrypted with the password set for this workbook. CSV exports are not encrypted.'
							: 'Saving writes this workbook without a password. Set a password here to encrypt Excel saves.',
					),
				),
				encrypt,
			);
			if (protectedSave) {
				const remove = primary(page, t('Remove password'), () => host.removePassword?.());
				remove.disabled = host.ctx.readOnly();
				left.append(remove);
			}
		}
		for (const [key, label] of PROPERTY_FIELDS) {
			const input = doc.createElement('input');
			input.type = 'text';
			input.value = workbook.properties[key] ?? '';
			input.readOnly = host.ctx.readOnly();
			input.addEventListener('change', () => host.setProperty(key, input.value));
			right.append(labelled(page, t(label), input));
		}
		const { cells, formulas } = countCells(workbook);
		right.append(customProperties(page, workbook));
		const stamp = (value: string | undefined) =>
			value ? new Date(value).toLocaleString(host.ctx.locale()) : '-';
		right.append(
			facts(page, [
				['Format', t(FORMAT_NAMES[workbook.format])],
				['Sheets', String(workbook.sheets.length)],
				['Cells with content', String(cells)],
				['Formulas', String(formulas)],
				['Defined names', String(workbook.definedNames.length)],
				['Created', stamp(workbook.properties.created)],
				['Last modified', stamp(workbook.properties.modified)],
			]),
		);
	}
	const layout = doc.createElement('div');
	layout.className = 'xve-backstage-columns';
	layout.append(left, right);
	content.replaceChildren(heading(page, t('Info')), layout);
}
