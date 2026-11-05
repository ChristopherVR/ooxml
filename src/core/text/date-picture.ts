/**
 * One date/time picture engine for every Office product.
 *
 * Word `\@` pictures and content-control date formats, Visio `{{...}}` field pictures and the
 * PowerPoint `a:fld` datetime types all describe a date with the same letters, so they share
 * this formatter instead of each carrying a copy. Case matters: `M` is the month, `m` the minute.
 *
 * Tokens: `yyyy` `yy` | `MMMM` `MMM` `MM` `M` | `dddd` `ddd` `dd` `d` (`EEEE` `EEE` are weekday
 * aliases) | `HH` `H` `hh` `h` | `mm` `m` | `ss` `s` | `tt` (AM/PM) `am/pm` `AM/PM` `A/P` `a/p`.
 * Text in single or double quotes is literal; any other character is copied as is.
 */

/** Options of {@link formatDatePicture}. */
export interface DatePictureOptions {
	/** Locale for month and weekday names (default `en-US`). */
	locale?: string;
	/** Read the date's UTC fields instead of its local ones (Visio stores wall-clock time as UTC). */
	utc?: boolean;
}

const TOKENS =
	/yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|EEEE|EEE|HH|H|hh|h|mm|m|ss|s|tt|am\/pm|AM\/PM|A\/P|a\/p|'[^']*'|"[^"]*"/g;

const pad = (value: number, length = 2): string => String(value).padStart(length, '0');

function name(
	date: Date,
	locale: string,
	utc: boolean,
	options: Intl.DateTimeFormatOptions,
): string {
	try {
		return new Intl.DateTimeFormat(locale, utc ? { ...options, timeZone: 'UTC' } : options).format(
			date,
		);
	} catch {
		return new Intl.DateTimeFormat('en-US', utc ? { ...options, timeZone: 'UTC' } : options).format(
			date,
		);
	}
}

/** Formats `date` with a date/time picture; see the module comment for the tokens. */
export function formatDatePicture(
	date: Date,
	picture: string,
	options: DatePictureOptions = {},
): string {
	const locale = options.locale ?? 'en-US';
	const utc = options.utc ?? false;
	const year = utc ? date.getUTCFullYear() : date.getFullYear();
	const month = utc ? date.getUTCMonth() : date.getMonth();
	const day = utc ? date.getUTCDate() : date.getDate();
	const hours = utc ? date.getUTCHours() : date.getHours();
	const minutes = utc ? date.getUTCMinutes() : date.getMinutes();
	const seconds = utc ? date.getUTCSeconds() : date.getSeconds();
	const hours12 = hours % 12 || 12;
	const label = (options: Intl.DateTimeFormatOptions) => name(date, locale, utc, options);
	return picture.replace(TOKENS, (token) => {
		switch (token) {
			case 'yyyy':
				return pad(year, 4);
			case 'yy':
				return pad(year % 100);
			case 'MMMM':
				return label({ month: 'long' });
			case 'MMM':
				return label({ month: 'short' });
			case 'MM':
				return pad(month + 1);
			case 'M':
				return String(month + 1);
			case 'dddd':
			case 'EEEE':
				return label({ weekday: 'long' });
			case 'ddd':
			case 'EEE':
				return label({ weekday: 'short' });
			case 'dd':
				return pad(day);
			case 'd':
				return String(day);
			case 'HH':
				return pad(hours);
			case 'H':
				return String(hours);
			case 'hh':
				return pad(hours12);
			case 'h':
				return String(hours12);
			case 'mm':
				return pad(minutes);
			case 'm':
				return String(minutes);
			case 'ss':
				return pad(seconds);
			case 's':
				return String(seconds);
			case 'tt':
			case 'AM/PM':
				return hours < 12 ? 'AM' : 'PM';
			case 'am/pm':
				return hours < 12 ? 'am' : 'pm';
			case 'A/P':
				return hours < 12 ? 'A' : 'P';
			case 'a/p':
				return hours < 12 ? 'a' : 'p';
			default:
				return token.slice(1, -1);
		}
	});
}

/** The English pattern of each predefined PresentationML datetime field type (ECMA-376 19.7.26). */
export const OOXML_DATETIME_PATTERNS: Readonly<Record<string, string>> = {
	datetime1: 'M/d/yyyy',
	datetime2: 'dddd, MMMM d, yyyy',
	datetime3: 'd MMMM yyyy',
	datetime4: 'MMMM d, yyyy',
	datetime5: 'dd-MMM-yy',
	datetime6: 'MMMM yy',
	datetime7: 'MMM-yy',
	datetime8: 'M/d/yyyy h:mm tt',
	datetime9: 'M/d/yyyy h:mm:ss tt',
	datetime10: 'H:mm',
	datetime11: 'H:mm:ss',
	datetime12: 'h:mm tt',
	datetime13: 'h:mm:ss tt',
};

/** Characters a format-encoded field type may carry: picture letters and plain separators. */
const ENCODED_SUFFIX = /^[yMdHhmsEt/.,:\- ]+$/;
const DATE_LETTER = /[yMdEhHs]/;

/**
 * The picture a field `type` stands for, or `undefined` when it names none.
 *
 * Predefined types (`datetime1`-`datetime13`, matched case-insensitively) map through
 * {@link OOXML_DATETIME_PATTERNS}. PowerPoint also writes the chosen format into the type itself
 * (`datetimeyyyy`, `datetimeMMMMyyyy`, `datetimeyyyy-MM-dd`): the suffix after `datetime` is the
 * picture, and is case-sensitive. A suffix with other letters (`datetimeFigureOut`) is not one.
 */
export function ooxmlDateTimeFieldPicture(type: string): string | undefined {
	const predefined = OOXML_DATETIME_PATTERNS[type.toLowerCase()];
	if (predefined !== undefined) {
		return predefined;
	}
	if (!/^datetime/i.test(type)) {
		return undefined;
	}
	const suffix = type.slice('datetime'.length);
	return ENCODED_SUFFIX.test(suffix) && DATE_LETTER.test(suffix) ? suffix : undefined;
}
