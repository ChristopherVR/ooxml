import { parseAddress } from '../address.js';
import type { DefinedName, FreezePane, SheetView } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

/** Freezes the top `rows` and left `cols` (both zero, or undefined, unfreezes). */
export function setFreeze(ctx: EditContext, s: number, freeze: FreezePane | undefined): void {
	const sheet = sheetAt(ctx.workbook, s);
	const next =
		freeze && (freeze.rows > 0 || freeze.cols > 0)
			? { rows: Math.max(0, freeze.rows), cols: Math.max(0, freeze.cols) }
			: undefined;
	ctx.run(
		next ? 'Freeze panes' : 'Unfreeze panes',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			if (next) sheet.view.freeze = next;
			else delete sheet.view.freeze;
		},
		{ sheet: s, structural: true },
	);
}

/** Patches the sheet view (zoom, gridlines, headers, selection, ...). */
export function setSheetView(ctx: EditContext, s: number, patch: Partial<SheetView>): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Sheet view',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const view = { ...sheet.view, ...structuredClone(patch) } as SheetView;
			for (const key of Object.keys(view) as (keyof SheetView)[])
				if (view[key] === undefined) delete view[key];
			if (patch.zoom !== undefined) view.zoom = Math.max(10, Math.min(400, Math.round(patch.zoom)));
			sheet.view = view;
		},
		{ sheet: s, structural: patch.zoom !== undefined || 'freeze' in patch },
	);
}

const NAME_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\?]*$/;

/** Why a defined name is not acceptable, or undefined when it is. */
export function validateDefinedName(name: string): string | undefined {
	if (!name || name.length > 255) return 'A name must have 1 to 255 characters.';
	if (!NAME_PATTERN.test(name))
		return 'A name must start with a letter, underscore or backslash and contain no spaces.';
	if (parseAddress(name) || /^[RrCc]$/.test(name) || /^R\d*C\d*$/i.test(name))
		return 'A name cannot look like a cell reference.';
	return undefined;
}

const sameName = (a: DefinedName, name: string, localSheet: number | undefined): boolean =>
	a.name.toLowerCase() === name.toLowerCase() && a.localSheet === localSheet;

/** Adds or replaces a defined name (matched by name, case-insensitively, and scope). */
export function setDefinedName(ctx: EditContext, name: DefinedName): void {
	const problem = name.name.startsWith('_xlnm.') ? undefined : validateDefinedName(name.name);
	if (problem) throw new Error(problem);
	if (name.localSheet !== undefined) sheetAt(ctx.workbook, name.localSheet);
	const entry = structuredClone(name);
	entry.formula = entry.formula.replace(/^=/, '');
	ctx.run('Define name', 'names', [{ kind: 'meta' }], () => {
		const { definedNames } = ctx.workbook;
		const index = definedNames.findIndex((n) => sameName(n, name.name, name.localSheet));
		if (index >= 0) definedNames[index] = entry;
		else definedNames.push(entry);
	});
}

export function deleteDefinedName(ctx: EditContext, name: string, localSheet?: number): void {
	const index = ctx.workbook.definedNames.findIndex((n) => sameName(n, name, localSheet));
	if (index < 0) return;
	ctx.run('Delete name', 'names', [{ kind: 'meta' }], () => {
		ctx.workbook.definedNames.splice(index, 1);
	});
}
