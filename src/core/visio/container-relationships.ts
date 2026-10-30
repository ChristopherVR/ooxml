/**
 * Visio's container relationships, as text. A container's Relationships cell lists its members
 * with `SUM(DEPENDSON(4, Sheet.N!SheetRef(), ...))`; each member's Relationships cell points back
 * with `SUM(DEPENDSON(1, Sheet.C!SheetRef(), ...))` (one reference per container it belongs to).
 * The formulas are never evaluated; they are read and written as text, and the cached value
 * stays 0 as in files Visio saves. The kind numbers are not checked against a Visio build.
 */
export const VISIO_RELATIONSHIP_MEMBERS = 4;
export const VISIO_RELATIONSHIP_MEMBER_OF = 1;
export const VISIO_RELATIONSHIP_CALLOUT_TARGET = 6;

export const SHEET_ID = /^[1-9]\d{0,9}$/;
export const MAX_RELATED_SHEETS = 1000;

/** Sheet IDs listed as `Sheet.N!SheetRef()` arguments of `DEPENDSON(kind, ...)`. Never evaluated. */
export function dependsOn(formula: string | undefined, kind: number): string[] {
	if (!formula || formula.length > 100_000) return [];
	const ids: string[] = [];
	for (const call of formula.matchAll(/DEPENDSON\(\s*(\d+)\s*,([^()]*(?:\([^()]*\)[^()]*)*)\)/gi))
		if (Number(call[1]) === kind)
			for (const ref of call[2]!.matchAll(/Sheet\.(\d+)!SheetRef\(\)/gi))
				if (SHEET_ID.test(ref[1]!) && ids.length < MAX_RELATED_SHEETS) ids.push(ref[1]!);
	return ids;
}
/** A comma separated list of sheet IDs, as the editor's private User row stores membership. */
export const sheetIdList = (value: string | undefined): string[] =>
	(value ?? '')
		.split(',')
		.map((id) => id.trim())
		.filter((id) => SHEET_ID.test(id))
		.slice(0, MAX_RELATED_SHEETS);

const RELATIONSHIPS_CALL = String.raw`DEPENDSON\(\s*\d+(?:\s*,\s*Sheet\.\d+!SheetRef\(\))*\s*\)`;
const RELATIONSHIPS_ONLY = new RegExp(
	String.raw`^SUM\(\s*${RELATIONSHIPS_CALL}(?:\s*,\s*${RELATIONSHIPS_CALL})*\s*\)$`,
	'i',
);
/**
 * Whether a Relationships formula only lists related sheets (`SUM(DEPENDSON(...), ...)`). Such a
 * formula registers a relationship and reads no cell of the sheets it names, so moving or resizing
 * one of them never changes its result.
 */
export const isRelationshipsOnly = (formula: string): boolean =>
	formula.length < 100_000 && RELATIONSHIPS_ONLY.test(formula.trim());

/** `SUM(DEPENDSON(kind, Sheet.a!SheetRef(), ...))` for the given sheet IDs. */
export function relationshipsFormula(kind: number, ids: readonly string[]): string {
	return `SUM(DEPENDSON(${kind}${ids.map((id) => `,Sheet.${id}!SheetRef()`).join('')}))`;
}
