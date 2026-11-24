/**
 * Marked content (ISO 32000-1 section 14.6): `BMC` / `BDC` ... `EMC` sequences, whose /MCID ties
 * what they draw to the structure tree, and whose /Artifact tag marks page furniture.
 */
import type { PdfDocument } from './document';
import { decodeTextString, isDict, isName, isNumber, type PdfDict, type PdfValue } from './objects';

export interface MarkedEntry {
	tag: string;
	mcid?: number;
	/** Alternate text given on the sequence itself (`/Alt`). */
	alt?: string;
	artifact?: true;
}

/** The entry a `BMC` (one operand) or `BDC` (tag and properties) opens. */
export async function markedEntry(
	document: PdfDocument,
	resources: PdfDict,
	operands: PdfValue[],
): Promise<MarkedEntry> {
	const tag = isName(operands[0]) ? operands[0].name : '';
	const entry: MarkedEntry = { tag };
	if (tag === 'Artifact') entry.artifact = true;
	let properties = operands[1];
	if (isName(properties))
		properties = (await document.dict(resources.get('Properties')))?.get(properties.name);
	const dict = await document.dict(properties);
	if (!dict || !isDict(dict)) return entry;
	const mcid = await document.resolve(dict.get('MCID'));
	if (isNumber(mcid) && mcid >= 0) entry.mcid = mcid;
	const alt = decodeTextString(await document.resolve(dict.get('Alt')));
	if (alt?.trim()) entry.alt = alt.trim();
	return entry;
}

/** The innermost marked-content id and alternate text, and whether any entry is an artifact. */
export function currentMark(stack: readonly MarkedEntry[]): {
	mcid?: number;
	alt?: string;
	artifact?: true;
} {
	const out: { mcid?: number; alt?: string; artifact?: true } = {};
	for (let index = stack.length - 1; index >= 0; index--) {
		const entry = stack[index]!;
		if (out.mcid === undefined && entry.mcid !== undefined) out.mcid = entry.mcid;
		if (out.alt === undefined && entry.alt !== undefined) out.alt = entry.alt;
		if (entry.artifact) out.artifact = true;
	}
	return out;
}
