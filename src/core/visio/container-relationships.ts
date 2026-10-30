/**
 * Visio's container relationships, as text. A container's Relationships cell lists its members
 * with `SUM(DEPENDSON(1, Sheet.N!SheetRef(), ...))`; each member's Relationships cell points back
 * with `DEPENDSON(4, Sheet.C!SheetRef(), ...)` (one reference per container it belongs to). A list
 * (a cross-functional flowchart's Swimlane List and Phase List) lists its members with kind 2 and
 * each member points back with kind 5; the two lists of one flowchart point at each other with
 * kind 11. The numbers were read from a cross-functional diagram Visio 16 built and confirmed
 * through its Container and List COM properties. The formulas are never evaluated; they are read
 * and written as text, and the cached value stays 0 as in files Visio saves. An empty container
 * keeps an empty `DEPENDSON(1)` term as Visio writes it. Files this editor wrote before
 * 2026-10-30 used 4 for the members and 1 for the back references; the reader still takes those
 * (containerMemberIds in shape-structure.ts).
 */
export const VISIO_RELATIONSHIP_MEMBERS = 1;
export const VISIO_RELATIONSHIP_MEMBER_OF = 4;
export const VISIO_RELATIONSHIP_LIST_MEMBERS = 2;
export const VISIO_RELATIONSHIP_LIST_MEMBER_OF = 5;
export const VISIO_RELATIONSHIP_LIST_SIBLING = 11;
/** Not confirmed against a Visio build. */
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

/** The terms of a relationships-only formula, in order; undefined for any other formula. */
export function relationTerms(
	formula: string | undefined,
): { kind: number; ids: string[] }[] | undefined {
	if (formula === undefined || !isRelationshipsOnly(formula)) return undefined;
	return [
		...formula.matchAll(/DEPENDSON\(\s*(\d+)\s*((?:,\s*Sheet\.\d+!SheetRef\(\)\s*)*)\)/gi),
	].map((call) => ({
		kind: Number(call[1]),
		ids: [...call[2]!.matchAll(/Sheet\.(\d+)!/gi)].map((ref) => ref[1]!),
	}));
}
/** `SUM(DEPENDSON(...), ...)` for terms; undefined when there are none. */
export function relationsFormula(
	terms: readonly { kind: number; ids: readonly string[] }[],
): string | undefined {
	if (!terms.length) return undefined;
	const call = ({ kind, ids }: { kind: number; ids: readonly string[] }) =>
		`DEPENDSON(${kind}${ids.map((id) => `,Sheet.${id}!SheetRef()`).join('')})`;
	return `SUM(${terms.map(call).join(',')})`;
}
/**
 * `formula` with the ids of one kind replaced and the other kinds kept in place. A term left
 * empty is dropped unless `keepEmpty` (a container's members). Undefined when nothing is left;
 * null when the formula holds more than relationships (the caller leaves it alone).
 */
export function withRelation(
	formula: string | undefined,
	kind: number,
	ids: readonly string[],
	keepEmpty = false,
): string | undefined | null {
	const terms = formula === undefined ? [] : relationTerms(formula);
	if (!terms) return null;
	const unique = [...new Set(ids)];
	const index = terms.findIndex((term) => term.kind === kind);
	if (index >= 0) terms[index] = { kind, ids: unique };
	else terms.push({ kind, ids: unique });
	return relationsFormula(terms.filter((term) => term.kind !== kind || unique.length || keepEmpty));
}
/** `formula` without the given sheets in any term (empty terms go, except an empty kind 1). */
export function withoutSheets(
	formula: string | undefined,
	removed: ReadonlySet<string>,
): string | undefined | null {
	const terms = relationTerms(formula);
	if (!terms) return null;
	return relationsFormula(
		terms
			.map((term) => ({ kind: term.kind, ids: term.ids.filter((id) => !removed.has(id)) }))
			.filter((term) => term.ids.length > 0 || term.kind === VISIO_RELATIONSHIP_MEMBERS),
	);
}
/** `SUM(DEPENDSON(kind, Sheet.a!SheetRef(), ...))` for the given sheet IDs. */
export function relationshipsFormula(kind: number, ids: readonly string[]): string {
	return `SUM(DEPENDSON(${kind}${ids.map((id) => `,Sheet.${id}!SheetRef()`).join('')}))`;
}
