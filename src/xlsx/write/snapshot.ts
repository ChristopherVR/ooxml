import { parseXml, type XmlElement } from '../../xml/index.js';
import type { Worksheet } from '../model.js';
import { snapshotKey } from '../read/worksheet.js';

function stable(value: unknown): unknown {
	if (value instanceof Map) return [...value.entries()].map(([k, v]) => [stable(k), stable(v)]);
	if (Array.isArray(value)) return value.map(stable);
	if (value && typeof value === 'object') {
		const out: Record<string, unknown> = {};
		for (const key of Object.keys(value).sort()) {
			const entry = (value as Record<string, unknown>)[key];
			if (entry !== undefined) out[key] = stable(entry);
		}
		return out;
	}
	return value;
}

/** Structural equality that ignores key order and `undefined` members. */
export const sameModel = (a: unknown, b: unknown): boolean =>
	JSON.stringify(stable(a)) === JSON.stringify(stable(b));

/** The source XML kept for a modelled element (see `SNAPSHOT_ELEMENTS`), parsed. */
export function snapshotElement(sheet: Worksheet, local: string): XmlElement | undefined {
	const xml = sheet.preserved.get(snapshotKey(local))?.[0];
	if (!xml) return undefined;
	try {
		return parseXml(xml, { label: 'XLSX worksheet' }).documentElement;
	} catch {
		return undefined;
	}
}

/** The raw snapshot string for an element. */
export const snapshotXml = (sheet: Worksheet, local: string): string | undefined =>
	sheet.preserved.get(snapshotKey(local))?.[0];
