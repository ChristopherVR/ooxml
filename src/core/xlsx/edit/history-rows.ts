// Undo snapshots for row dimension edits (heights, hiding, outline levels). They record only the
// `rowInfo` entries of the rows an edit names, not the sheet or the rest of its row metadata.
import type { RowInfo, Worksheet } from '../model';

/** Each named row's metadata, cloned; `undefined` where the row has none. */
export type RowInfoEntries = [row: number, info: RowInfo | undefined][];

export function captureRowInfo(sheet: Worksheet, rows: readonly number[]): RowInfoEntries {
	return [...new Set(rows)].map((row) => {
		const info = sheet.rowInfo.get(row);
		return [row, info && structuredClone(info)];
	});
}

/** Puts recorded row metadata back into the sheet's own `rowInfo` map. */
export function restoreRowInfo(sheet: Worksheet, entries: RowInfoEntries): void {
	for (const [row, info] of entries)
		if (info) sheet.rowInfo.set(row, structuredClone(info));
		else sheet.rowInfo.delete(row);
}
