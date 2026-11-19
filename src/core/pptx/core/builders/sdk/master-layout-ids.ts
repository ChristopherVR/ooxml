/**
 * Slide master and slide layout ids (`p:sldMasterId` in presentation.xml, `p:sldLayoutId` in each
 * master) share one id space that starts at 2147483648, so every new one must be above all of
 * them, not only above its own kind or its own master's list. A collision makes the package
 * ambiguous: slide paste, for one, refuses such a destination.
 *
 * @module sdk/master-layout-ids
 */
import type JSZip from 'jszip';

const FIRST_ID = 2147483648;
const ID_PATTERN = /<(?:\w+:)?sld(?:Master|Layout)Id\b[^>]*?\sid="(\d+)"/g;

function isIdPart(path: string): boolean {
	return path === 'ppt/presentation.xml' || /^ppt\/slideMasters\/slideMaster\d+\.xml$/u.test(path);
}

/** The lowest master or layout id above every one the package already uses. */
export async function nextMasterLayoutId(zip: JSZip): Promise<number> {
	let next = FIRST_ID;
	for (const path of Object.keys(zip.files).filter(isIdPart)) {
		const xml = (await zip.file(path)?.async('string')) ?? '';
		for (const match of xml.matchAll(ID_PATTERN)) {
			next = Math.max(next, Number(match[1]) + 1);
		}
	}
	return next;
}

/** Give every layout listed in a master's XML a fresh id, counting up from `first`. */
export function renumberLayoutIds(masterXml: string, first: number): string {
	let id = first;
	return masterXml.replace(
		/(<(?:\w+:)?sldLayoutId\b[^>]*?\sid=")\d+(")/g,
		(_match, head: string, tail: string) => `${head}${id++}${tail}`,
	);
}
