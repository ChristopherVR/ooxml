// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { formatDatePicture } from '../text/date-picture';

// Word date/time field results: the `\@` date-time picture switch (e.g. `\@ "MMMM d, yyyy"`).

/** The `\@` picture of a field instruction, if any (quotes optional). */
export function datePicture(instr: string): string | undefined {
	const match = /\\@\s*(?:"([^"]*)"|(\S+))/.exec(instr);
	return match ? (match[1] ?? match[2]) : undefined;
}

/**
 * Formats `date` with a Word date-time picture. Month and day names follow `locale`; text in
 * quotes is literal. Unrecognized characters are copied as-is, as Word does.
 */
export function formatWordDate(date: Date, picture: string, locale = 'en-US'): string {
	return formatDatePicture(date, picture, { locale });
}

/** DATE and TIME results as Word shows them when it updates the field (default pictures included). */
export function dateFieldResult(
	name: 'DATE' | 'TIME',
	instr: string,
	now: Date,
	locale?: string,
): string {
	const picture = datePicture(instr) ?? (name === 'DATE' ? 'M/d/yyyy' : 'h:mm am/pm');
	return formatWordDate(now, picture, locale);
}
