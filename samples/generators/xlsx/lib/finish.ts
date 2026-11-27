/**
 * Calculating and saving a sample workbook: every formula gets its cached value from the core's
 * formula engine, and the package is written by `saveXlsx` with fixed document properties.
 */
import { createEditSession } from '../../../../src/core/xlsx/edit/index';
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import { formatAddress } from '../../../../src/core/xlsx/address';
import { isCellError } from '../../../../src/core/xlsx/model';
import { saveXlsx } from '../../../../src/core/xlsx/write/index';
import { FONTS } from '../../lib/fiction';
import { patchPackage, readPartText, xmlEscape } from '../../lib/package';
import { SAMPLE_AUTHOR, SAMPLE_DATE_ISO } from '../../lib/types';
import type { Book } from './book';
import { fitNotes } from './notes';

export interface BookInfo {
	title: string;
	subject: string;
	description: string;
	keywords: string;
	company: string;
}

/**
 * An edit session over the finished cells with every formula calculated. Features the core
 * builds from calculated data (PivotTables, slicers, sparklines, charts) are added through it.
 */
export function calculate(book: Book): EditSession {
	// Row heights are designed by hand, and nothing here is undone.
	const session = createEditSession(book.wb, { autoRowHeight: false, historyLimit: 1 });
	session.calculateNow({ full: true });
	return session;
}

/** Formula cells that hold an error: a sample must not ship one by accident. */
function formulaErrors(book: Book): string[] {
	const found: string[] = [];
	for (const sheet of book.wb.sheets)
		for (const [row, cells] of sheet.rows)
			for (const [col, cell] of cells)
				if (cell.formula !== undefined && isCellError(cell.value))
					found.push(`${sheet.name}!${formatAddress({ row, col })} ${cell.value.error}`);
	return found;
}

/**
 * Finishes the theme part. The core writes its palette and fonts into a theme that is still
 * called Office (what Excel would then show for these custom colours) and leaves the East Asian
 * and complex script slots empty, so both are filled in here.
 */
async function brandTheme(bytes: Uint8Array, company: string): Promise<Uint8Array> {
	const part = 'xl/theme/theme1.xml';
	const xml = await readPartText(bytes, part);
	if (!xml) throw new Error('The workbook has no theme part');
	const name = xmlEscape(company);
	const slots = (font: string) => `<a:ea typeface="${font}"/><a:cs typeface="${font}"/>`;
	const empty = slots('');
	return patchPackage(bytes, {
		[part]: xml
			.replace('name="Office Theme"', `name="${name}"`)
			.replace('<a:clrScheme name="Office">', `<a:clrScheme name="${name}">`)
			.replace(
				'<a:fontScheme name="Office">',
				`<a:fontScheme name="${FONTS.display} and ${FONTS.body}">`,
			)
			.replace('<a:fmtScheme name="Office">', `<a:fmtScheme name="${name}">`)
			// The major font comes first in the part, then the minor font.
			.replace(empty, slots(FONTS.display))
			.replace(empty, slots(FONTS.body)),
	});
}

/**
 * Recalculates, checks the result and writes the package. `patch` rewrites parts after the save,
 * for the few details the core writer does not cover yet.
 */
export async function finish(
	book: Book,
	session: EditSession,
	info: BookInfo,
	patch?: (bytes: Uint8Array) => Promise<Uint8Array>,
): Promise<Uint8Array> {
	session.calculateNow({ full: true });
	const errors = formulaErrors(book);
	if (errors.length)
		throw new Error(`${info.title}: formula errors: ${errors.slice(0, 12).join(', ')}`);
	if (book.wb.warnings.length) throw new Error(`${info.title}: ${book.wb.warnings.join('; ')}`);
	book.wb.properties = {
		...book.wb.properties,
		title: info.title,
		subject: info.subject,
		description: info.description,
		keywords: info.keywords,
		company: info.company,
		creator: SAMPLE_AUTHOR,
		lastModifiedBy: SAMPLE_AUTHOR,
		created: SAMPLE_DATE_ISO,
		modified: SAMPLE_DATE_ISO,
		application: 'ooxml-core',
	};
	const bytes = await brandTheme(await fitNotes(await saveXlsx(book.wb), book), info.company);
	return patch ? patch(bytes) : bytes;
}
