/** Microsoft Cell_Type F markers are distinct from executable formulas.
 * No Formula locally deletes/blocks a formula; Inh delegates to inheritance.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/cell-element-geometry-sectionvisio-xml
 * Keep exact spelling: an unknown string must never become an inert marker.
 */
export function executableCellFormula(source: string | undefined): string | undefined {
	return !source || source === 'No Formula' || source === 'Inh' ? undefined : source;
}

/** Literal double-click handlers execute only on user activation, never on a transform. */
export function inertDoubleClickFormula(cell: string, source: string): boolean {
	return cell === 'EventDblClick' && /^\s*=?\s*(OPENTEXTWIN|DEFAULTEVENT)\(\s*\)\s*$/i.test(source);
}
