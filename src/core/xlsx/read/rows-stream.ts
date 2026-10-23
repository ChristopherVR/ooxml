// The rows of a worksheet part read in pieces: the part is inflated chunk by chunk under the zip
// guard and decoded as it arrives, and its cell data is cut before the last row start that has
// arrived, so each piece holds whole rows for the light reader, which builds one row at a time.
// Where the light reader declines a piece (markup it leaves to the XML parser, or a comment or
// CDATA section a cut went through), the whole part is read the way `loadXlsx` reads it and the
// rows go on from where they had got to.
import type JSZip from 'jszip';
import { streamEntryGuarded } from '../../opc/index';
import { LiteDeclinedError, elements, liteChildElements, type XmlElement } from '../../xml/index';
import type { Cell } from '../model';
import { createWorksheet } from '../workbook';
import { parseDataXml, zipLimitsFor, type XlsxLoadLimits } from './limits';
import { sheetDataReader, type CellContext } from './sheet-data';
import { sheetDataStart, splitSheetData, type SheetDataStart } from './sheet-data-source';

/** How much of a part is read looking for `<sheetData>` before the part is read whole. */
const MAX_HEAD = 1 << 20;

/** A row's index and its cells as `loadXlsx` stores them. */
export type SheetRow = [row: number, cells: Map<number, Cell>];

/**
 * The rows of the worksheet part in `entry`, each as soon as it is read. `alreadyRead` is what
 * the reader inflated before (for the total limit).
 */
export async function* sheetRows(
	ctx: CellContext,
	entry: JSZip.JSZipObject,
	alreadyRead: number,
	limits: Required<XlsxLoadLimits>,
): AsyncGenerator<SheetRow, void, undefined> {
	const sheet = createWorksheet('', 1);
	const reader = sheetDataReader(ctx, sheet);
	let consumed = 0;
	/** Reads the items after the first `skip`, handing out the rows each one adds. */
	function* take(items: Iterable<XmlElement>, skip: number): Generator<SheetRow> {
		let index = 0;
		for (const item of items) {
			if (index++ < skip) continue;
			reader.read(item);
			consumed++;
			sheet.rowInfo.clear();
			delete sheet.dataTables;
			if (!sheet.rows.size) continue;
			const rows = [...sheet.rows].sort((a, b) => a[0] - b[0]);
			sheet.rows.clear();
			yield* rows;
		}
	}
	try {
		for await (const piece of pieces(entry, alreadyRead, limits))
			yield* take(liteChildElements(piece.xml, piece.start.namespaces), 0);
		return;
	} catch (error) {
		if (!(error instanceof LiteDeclinedError)) throw error;
	}
	const xml = new TextDecoder('utf-8').decode(await entry.async('uint8array'));
	const split = splitSheetData(xml);
	if (split)
		try {
			yield* take(split.rows, consumed);
			return;
		} catch (error) {
			if (!(error instanceof LiteDeclinedError)) throw error;
		}
	const root = parseDataXml(xml, 'XLSX worksheet', limits.maxXmlElements).documentElement;
	const data = elements(root).find((node) => node.localName === 'sheetData');
	yield* take(data ? elements(data) : [], consumed);
}

/**
 * The cell data of the part in pieces that end before a row start, with where it started.
 * Throws a `LiteDeclinedError` when the cell data cannot be found this way.
 */
async function* pieces(
	entry: JSZip.JSZipObject,
	alreadyRead: number,
	limits: Required<XlsxLoadLimits>,
): AsyncGenerator<{ xml: string; start: SheetDataStart }, void, undefined> {
	const decoder = new TextDecoder('utf-8');
	let start: SheetDataStart | undefined;
	let row = '';
	let buffer = '';
	const chunks = streamEntryGuarded(entry, alreadyRead, zipLimitsFor(limits));
	try {
		for (let ended = false; !ended;) {
			const next = await chunks.next();
			ended = next.done === true;
			buffer += next.done ? decoder.decode() : decoder.decode(next.value, { stream: true });
			if (!start) {
				const found = sheetDataStart(buffer);
				if (found === false || (!found && (ended || buffer.length > MAX_HEAD)))
					throw new LiteDeclinedError();
				if (!found) continue;
				start = found;
				row = `<${found.tagName.slice(0, -'sheetData'.length)}row`;
				buffer = buffer.slice(found.contentStart);
			}
			const close = buffer.indexOf(`</${start.tagName}`);
			if (close >= 0) {
				yield { xml: buffer.slice(0, close), start };
				return;
			}
			const cut = lastRowStart(buffer, row);
			if (cut > 0) {
				yield { xml: buffer.slice(0, cut), start };
				buffer = buffer.slice(cut);
			}
		}
	} finally {
		// Stops inflating when the cell data has ended or the reader stopped early.
		await chunks.return();
	}
	throw new LiteDeclinedError();
}

/** Where the last `<row` start tag (`row`, with its prefix) in `xml` begins, past the start. */
function lastRowStart(xml: string, row: string): number {
	for (let at = xml.lastIndexOf(row); at > 0; at = xml.lastIndexOf(row, at - 1)) {
		const next = xml.charCodeAt(at + row.length);
		// A space, `>` or `/` ends the name (`<rows` is another element).
		if (next === 32 || next === 9 || next === 10 || next === 13 || next === 62 || next === 47)
			return at;
	}
	return -1;
}
