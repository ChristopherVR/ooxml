import { parseActionSteps } from './action-formula';
import { isRelationshipsOnly } from './container-relationships';

/** Microsoft Cell_Type F markers are distinct from executable formulas.
 * No Formula locally deletes/blocks a formula; Inh delegates to inheritance.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/cell-element-geometry-sectionvisio-xml
 * Keep exact spelling: an unknown string must never become an inert marker.
 */
export function executableCellFormula(source: string | undefined): string | undefined {
	return !source || source === 'No Formula' || source === 'Inh' ? undefined : source;
}

/**
 * Formulas that never run on a transform: literal double-click handlers (they execute only on
 * user activation) and the Relationships cell of a container or member, which is a text list of
 * `SUM(DEPENDSON(kind, Sheet.N!SheetRef(), ...))` links whose cached value stays 0, and an
 * Action cell made of local `SETF(GetRef(cell), value)` steps, which runs only when the user
 * picks the action (the built-in masters' Set to Default Size).
 */
export function inertDoubleClickFormula(cell: string, source: string): boolean {
	return (
		(cell === 'EventDblClick' && /^\s*=?\s*(OPENTEXTWIN|DEFAULTEVENT)\(\s*\)\s*$/i.test(source)) ||
		(cell === 'Relationships' && isRelationshipsOnly(source)) ||
		(cell === 'Action' && 'steps' in parseActionSteps(source))
	);
}

/** Shape Data uses the XML section Property and the formula prefix Prop. */
export const shapeDataFormulaSection = (name: string): string =>
	name === 'Property' ? 'Prop' : name;
