import type { TextRun } from './model';
import { onOffAttribute } from './simple-types';
import { WORD_NS, type XmlElement } from './xml';

export type FieldFlags = NonNullable<TextRun['fieldFlags']>;

/** Absence stays distinct from an explicit false; invalid ST_OnOff tokens remain unset. */
export function parseFieldFlags(element: XmlElement): FieldFlags | undefined {
	const locked = onOffAttribute(element, 'fldLock');
	const dirty = onOffAttribute(element, 'dirty');
	if (locked === undefined && dirty === undefined) return undefined;
	return { ...(locked !== undefined && { locked }), ...(dirty !== undefined && { dirty }) };
}

/** Field structure owns these flags, independently of run formatting and text revisions. */
export function writeFieldFlags(element: XmlElement, flags: FieldFlags | undefined): void {
	for (const [name, value] of [
		['fldLock', flags?.locked],
		['dirty', flags?.dirty],
	] as const) {
		if (value === undefined) element.removeAttributeNS(WORD_NS, name);
		else element.setAttributeNS(WORD_NS, `w:${name}`, value ? 'true' : 'false');
	}
}
