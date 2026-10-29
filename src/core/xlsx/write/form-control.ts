// Keeps the cached state of the form controls (`formControlPr` `checked`, `sel`, `val`) in step
// with their linked cells, which are the source of truth in the editor. Parts whose state already
// matches stay byte-identical.
import type { FormControlObject, Workbook, Worksheet } from '../model';
import { controlStateIn } from '../layout/form-control-link';
import type { PackageWriter } from './package-writer';

/** Sets (or, with `null`, removes) one attribute of the root element of a small XML part. */
function setRootAttribute(xml: string, name: string, value: string | null): string {
	return xml.replace(/<formControlPr\b[^>]*?(\/?)>/, (tag, close: string) => {
		const pattern = new RegExp(`\\s${name}="[^"]*"`);
		const without = tag.replace(pattern, '');
		if (value === null) return without;
		const insert = ` ${name}="${value}"`;
		return without.replace(/\s*\/?>$/, `${insert}${close}>`);
	});
}

function patchedPart(xml: string, wanted: Record<string, string | null>): string {
	let out = xml;
	for (const [name, value] of Object.entries(wanted)) {
		const current = new RegExp(`\\s${name}="([^"]*)"`).exec(out)?.[1] ?? null;
		if (current !== value) out = setRootAttribute(out, name, value);
	}
	return out;
}

export function refreshFormControlParts(
	writer: PackageWriter,
	workbook: Workbook,
	sheet: Worksheet,
): void {
	const sheetIndex = workbook.sheets.indexOf(sheet);
	const controls = sheet.drawings.filter(
		(object): object is FormControlObject => object.kind === 'formControl',
	);
	for (const control of controls) {
		const part = control.partName;
		const source = part ? writer.source?.text(part) : undefined;
		if (!part || !source || !writer.has(part)) continue;
		const { state, link } = controlStateIn(workbook, sheetIndex, controls, control);
		if (!link) continue;
		const wanted: Record<string, string | null> = {};
		if (control.controlType === 'checkbox' || control.controlType === 'radio')
			wanted.checked =
				state.checked === 'checked' ? 'Checked' : state.checked === 'mixed' ? 'Mixed' : null;
		else if (control.controlType === 'dropdown' || control.controlType === 'list')
			wanted.sel = String(state.selected);
		else if (control.controlType === 'spin' || control.controlType === 'scroll')
			wanted.val = String(state.value);
		const next = patchedPart(source, wanted);
		if (next !== source) writer.replace(part, next);
	}
}
