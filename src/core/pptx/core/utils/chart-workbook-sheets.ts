// Which sheet of an embedded workbook a save that moves a chart's data writes, and which of the
// session's chart data sheets it removes, so saving any state again converges on one workbook.
import type JSZip from 'jszip';
import type { AppendChartWorkbookSheetOptions } from '../../../chart';
import { parseRelationships, resolvePartPath } from '../../../opc/relationships';

/** What a session knows of the chart data sheets of its embedded workbooks. Names are lower case. */
export interface ChartDataSheets {
	/** The sheets the session allocated for chart data, by embedded workbook path. */
	allocated: Map<string, Set<string>>;
	/** The sheets read by the saved states of a chart part that undo can still bring back. */
	held(chartPartPath: string): ReadonlySet<string>;
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** The sheets the formulas (`c:f`, `cx:f`, `c15:f`) of a chart part name, lower case. */
export function chartPartSheets(xml: string): Set<string> {
	const sheets = new Set<string>();
	for (const [, text] of xml.matchAll(/<(?:[\w.-]+:)?f(?:\s[^>]*)?>([^<]*)</gu)) {
		const formula = text!.replace(
			/&(amp|lt|gt|quot|apos);/gu,
			(_, name: string) => ENTITIES[name]!,
		);
		for (const [, quoted, bare] of formula.matchAll(
			/(?:'((?:[^']|'')+)'|([^\s!(),'=+\-*/&^<>]+))!/gu,
		))
			sheets.add((quoted?.replace(/''/gu, "'") ?? bare!).toLowerCase());
	}
	return sheets;
}

/** The sheets of `workbook` that parts of the package other than `part` read. */
async function sheetsReadElsewhere(zip: JSZip, workbook: string, part: string) {
	const sheets = new Set<string>();
	for (const relsPath of Object.keys(zip.files)) {
		const match = /^(.*?)_rels\/([^/]+)\.rels$/u.exec(relsPath);
		const source = match && `${match[1]}${match[2]}`;
		if (!source || source === part || !zip.file(source)) continue;
		const rels = parseRelationships(await zip.file(relsPath)!.async('string'));
		const reads = [...rels.values()].some(
			(rel) =>
				rel.mode !== 'External' &&
				rel.type.endsWith('/package') &&
				resolvePartPath(source, rel.target) === workbook,
		);
		if (reads)
			for (const sheet of chartPartSheets(await zip.file(source)!.async('string')))
				sheets.add(sheet);
	}
	return sheets;
}

/**
 * Where the chart part `part` about to move its data in `workbook` puts it: over the sheet this
 * session allocated that the part as it is now reads, when no other part reads it, or else on a new
 * one. The session's other chart data sheets that no part and no saved state undo can bring back
 * reads are removed. A sheet the workbook had when it was loaded is never written over or removed.
 * Call `written` with the sheet the data went to once the workbook is saved.
 */
export async function placeChartData(
	zip: JSZip,
	workbook: string,
	part: string,
	sheets: ChartDataSheets,
): Promise<{
	options: Pick<AppendChartWorkbookSheetOptions, 'reuse' | 'drop'>;
	written(sheetName: string): void;
}> {
	const ours = [...(sheets.allocated.get(workbook) ?? [])];
	const elsewhere = await sheetsReadElsewhere(zip, workbook, part);
	const held = sheets.held(part);
	const replaced = chartPartSheets((await zip.file(part)?.async('string')) ?? '');
	const reuse = ours.filter((sheet) => replaced.has(sheet) && !elsewhere.has(sheet));
	const read = (sheet: string) => elsewhere.has(sheet) || held.has(sheet);
	return {
		options: {
			...(reuse.length === 1 ? { reuse: reuse[0] } : {}),
			drop: ours.filter((sheet) => !read(sheet)),
		},
		written: (sheetName) =>
			sheets.allocated.set(workbook, new Set([sheetName.toLowerCase(), ...ours.filter(read)])),
	};
}
