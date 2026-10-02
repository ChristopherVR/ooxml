import type { PageSetup, PrintOptions } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

/** A page setup patch; an `undefined` entry clears that setting. */
export type PageSetupPatch = { [K in keyof PageSetup]?: PageSetup[K] | undefined };

/** Drops `undefined` entries so patches can clear a setting. */
function compact<T extends object>(value: T): T {
	for (const key of Object.keys(value) as (keyof T)[])
		if (value[key] === undefined) delete value[key];
	return value;
}

/** Patches the page setup (orientation, paper, scaling, margins, print area, header/footer). */
export function setPageSetup(ctx: EditContext, s: number, patch: PageSetupPatch): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Page setup',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const next = compact({ ...sheet.pageSetup, ...structuredClone(patch) }) as PageSetup;
			if (Object.keys(next).length) sheet.pageSetup = next;
			else delete sheet.pageSetup;
		},
		{ sheet: s },
	);
}

/** Patches the print options (gridlines, headings, centring); `false` or undefined clears one. */
export function setPrintOptions(ctx: EditContext, s: number, patch: Partial<PrintOptions>): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Print options',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const next: PrintOptions = { ...sheet.printOptions, ...patch };
			for (const key of Object.keys(next) as (keyof PrintOptions)[])
				if (!next[key]) delete next[key];
			if (Object.keys(next).length) sheet.printOptions = next;
			else delete sheet.printOptions;
			// The model now owns the element; never write a stale verbatim copy as well.
			sheet.preserved.delete('printOptions');
		},
		{ sheet: s },
	);
}
