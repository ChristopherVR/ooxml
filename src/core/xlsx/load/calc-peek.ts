// Whether a file asks to be calculated on open, read from its workbook part alone, so an editor
// can start that calculation elsewhere while it reads the whole file.
import { strFromU8, unzipSync } from 'fflate';

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
const WORKBOOK_PART = 'xl/workbook.xml';
/** A larger workbook part is not inflated just for a hint. */
const MAX_PART_BYTES = 4 * 1024 * 1024;
const CALC_PR = /<(?:[\w.-]+:)?calcPr\b([^<>]*)>/u;

/**
 * Whether an .xlsx or .xlsm file asks for a full calculation on open (`calcPr fullCalcOnLoad`)
 * and calculates automatically, judged from its workbook part alone (`fflate`'s filter reads the
 * central directory and inflates nothing else). False for anything that is not such a package
 * (encrypted, legacy or CSV) or cannot be read: it is a hint, and what `needsCalculationOnOpen`
 * says of the workbook the reader returns decides.
 */
export function asksForCalculationOnOpen(input: Uint8Array | ArrayBuffer): boolean {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (!ZIP_MAGIC.every((value, index) => bytes[index] === value)) return false;
	try {
		const parts = unzipSync(bytes, {
			filter: (file) => file.name === WORKBOOK_PART && file.originalSize <= MAX_PART_BYTES,
		});
		const part = parts[WORKBOOK_PART];
		const attrs = part ? (CALC_PR.exec(strFromU8(part))?.[1] ?? '') : '';
		return (
			/\bfullCalcOnLoad\s*=\s*["'](?:1|true)["']/u.test(attrs) &&
			!/\bcalcMode\s*=\s*["']manual["']/u.test(attrs)
		);
	} catch {
		return false;
	}
}
