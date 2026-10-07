/** Transform formula syntax while preserving double-quoted string literals exactly. */
export function mapVisioFormulaSyntax(
	source: string,
	transform: (segment: string) => string,
): string {
	return source
		.split(/("(?:[^"]|"")*")/g)
		.map((part, index) => (index % 2 ? part : transform(part)))
		.join('');
}
/** Mask strings before scanning executable function/reference syntax. */
export const unquotedVisioFormula = (source: string) =>
	source
		.split(/("(?:[^"]|"")*")/g)
		.map((part, index) => (index % 2 ? ' ' : part))
		.join('');
