// Built-in number formats as the package reader and writer need them: thin wrappers over the
// number-format module's table so there is a single list.
import { BUILTIN_NUMBER_FORMATS, builtinFormatId } from '../numfmt/builtins.js';

/** The built-in number formats by id (en-US forms); the numfmt module's table. */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = BUILTIN_NUMBER_FORMATS;

/**
 * The built-in id of a format code when the code is exactly that built-in's text (spelling
 * variants keep a declared `numFmt` so their text survives a save).
 */
export function builtinFormatIdOf(code: string): number | undefined {
	const id = builtinFormatId(code);
	return id !== undefined && BUILTIN_NUMBER_FORMATS[id] === code ? id : undefined;
}

/** Locale-dependent ids (27-36, 50-81: CJK and Thai date forms) fall back to a date format. */
export function builtinFormat(id: number): string | undefined {
	const known = BUILTIN_NUMBER_FORMATS[id];
	if (known !== undefined) return known;
	if ((id >= 27 && id <= 36) || (id >= 50 && id <= 81)) return 'm/d/yyyy';
	return undefined;
}
