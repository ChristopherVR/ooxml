/**
 * Sizes legacy note boxes to their text. The core writer gives every note the same small box,
 * which clips anything longer than a sentence, so the VML drawing is patched after the save.
 */
import type { Worksheet } from '../../../../src/core/xlsx/model';
import { patchPackage, readPackage } from '../../lib/package';
import type { Book } from './book';

const WIDTH = 300;
/** Characters of 10 point body text that fit one line of the box. */
const LINE_CHARS = 48;
const LINE_HEIGHT = 17;

const columnPx = (sheet: Worksheet, col: number): number => {
	const info = sheet.columns.find((entry) => col >= entry.min && col <= entry.max);
	return info?.width === undefined ? 64 : Math.round(info.width * 7 + 5);
};
const rowPx = (sheet: Worksheet, row: number): number =>
	Math.round(((sheet.rowInfo.get(row)?.height ?? sheet.defaultRowHeight) * 4) / 3);

/** The cell and offset reached by walking `distance` pixels from the start of `from`. */
function walk(size: (index: number) => number, from: number, distance: number): [number, number] {
	let index = from;
	let left = distance;
	while (left >= size(index)) {
		left -= size(index);
		index++;
	}
	return [index, left];
}

function boxFor(
	sheet: Worksheet,
	row: number,
	col: number,
	text: string,
): { anchor: string; style: string } {
	const lines = text
		.split('\n')
		.reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / LINE_CHARS)), 0);
	const height = lines * LINE_HEIGHT + 14;
	const [right, rightOffset] = walk((c) => columnPx(sheet, c), col + 1, 12 + WIDTH);
	const top = Math.max(0, row - 1);
	const [bottom, bottomOffset] = walk((r) => rowPx(sheet, r), top, 6 + height);
	return {
		anchor: `${col + 1}, 12, ${top}, 6, ${right}, ${rightOffset}, ${bottom}, ${bottomOffset}`,
		style: `width:${WIDTH * 0.75}pt;height:${height * 0.75}pt`,
	};
}

const decoder = new TextDecoder();

/** The package with every legacy note box of `book` resized; threaded comments are left alone. */
export async function fitNotes(bytes: Uint8Array, book: Book): Promise<Uint8Array> {
	const parts = await readPackage(bytes);
	const changes: Record<string, string> = {};
	book.wb.sheets.forEach((sheet, index) => {
		if (!sheet.comments.length) return;
		const rels = parts.get(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`);
		const target =
			rels && /Target="\.\.\/drawings\/(vmlDrawing\d+\.vml)"/.exec(decoder.decode(rels));
		const name = target ? `xl/drawings/${target[1]}` : undefined;
		const vml = name ? parts.get(name) : undefined;
		if (!name || !vml) throw new Error(`No VML drawing for the notes of ${sheet.name}`);
		// The writer lists the shapes in reading order, as it does the comments.
		const sorted = [...sheet.comments].sort(
			(a, b) => a.address.row - b.address.row || a.address.col - b.address.col,
		);
		let at = 0;
		changes[name] = decoder.decode(vml).replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) => {
			const comment = sorted[at++];
			if (!comment || comment.replies) return shape;
			const box = boxFor(sheet, comment.address.row, comment.address.col, comment.text);
			return shape
				.replace(/width:[\d.]+pt;height:[\d.]+pt/, box.style)
				.replace(/<x:Anchor>[^<]*<\/x:Anchor>/, `<x:Anchor>${box.anchor}</x:Anchor>`);
		});
	});
	return Object.keys(changes).length ? patchPackage(bytes, changes) : bytes;
}
