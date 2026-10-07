import type { TextRun } from './model';

/** A locked field retains its saved result during automatic and explicitly requested updates. */
export function isFieldLocked(run: Pick<TextRun, 'fieldFlags'> | undefined): boolean {
	return run?.fieldFlags?.locked === true;
}

/** Structural field boundaries keep explicit false distinct from omitted flags. */
export function sameFieldFlags(a: TextRun['fieldFlags'], b: TextRun['fieldFlags']): boolean {
	return a?.locked === b?.locked && a?.dirty === b?.dirty;
}
