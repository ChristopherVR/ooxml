import type { WorkbookProperties } from '../model.js';
import type { EditContext } from './context.js';

/** A change to the document properties: a value sets the field, `null` clears it. */
export type DocumentPropertiesPatch = {
	[K in keyof WorkbookProperties]?: WorkbookProperties[K] | null;
};

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/**
 * Applies `patch` to the workbook's document properties (core, extended and custom) as one
 * undo step. The properties object is replaced, never mutated, so history snapshots hold the
 * prior object. `custom` replaces the whole list; `custom: []` removes `docProps/custom.xml`.
 * Returns whether anything changed.
 */
export function setDocumentProperties(ctx: EditContext, patch: DocumentPropertiesPatch): boolean {
	const prior = ctx.workbook.properties;
	const next: Record<string, unknown> = { ...prior };
	for (const [key, value] of Object.entries(patch)) {
		if (value === undefined) continue;
		if (value === null) delete next[key];
		else next[key] = structuredClone(value);
	}
	if (same(next, prior)) return false;
	ctx.run('Document properties', 'view', [{ kind: 'meta' }], () => {
		ctx.workbook.properties = next as WorkbookProperties;
	});
	return true;
}
