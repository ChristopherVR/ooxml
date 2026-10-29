// The documents of `scripts/docx-word-table-probe.ps1`, rebuilt in TypeScript so the layout can be
// compared with what Word measured (`__fixtures__/word-table/table-probe.json`). Keep the two in
// step: every group, label, border and height here is the script's.
import { packProbe, para } from './word-spacing-docs';

const cell = (paras: string, tcPr = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${paras}</w:tc>`;
const row = (cells: string, trPr = '') => `<w:tr><w:trPr>${trPr}</w:trPr>${cells}</w:tr>`;
function tbl(rows: string, cols: number, border: string): string {
	const [val, sz] = border.split(':');
	const b = border
		? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
				.map((side) => `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="auto"/>`)
				.join('')}</w:tblBorders>`
		: '';
	return `<w:tbl><w:tblPr><w:tblW w:w="${3000 * cols}" w:type="dxa"/>${b}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(cols)}</w:tblGrid>${rows}</w:tbl>`;
}
const lines = (id: string, n: number, after = '0') =>
	Array.from({ length: n }, (_, i) => para(`${id}l${i + 1}`, `w:after="${after}"`)).join('');

export type TableGroup = 'vmerge' | 'borders' | 'rowh';

/** Document name to its group and body XML. */
export function tableBodies(): Record<string, { group: TableGroup; body: string }> {
	const out: Record<string, { group: TableGroup; body: string }> = {};
	for (const [bk, border] of [
		['n', ''],
		['s', 'single:4'],
	] as const)
		for (const col of [0, 1])
			for (const span of [2, 3])
				for (const n of [1, 2, 4, 7]) {
					const id = `vm${bk}_c${col}_s${span}_n${n}`;
					let rows = '';
					for (let r = 1; r <= span + 1; r++) {
						const cells: string[] = [];
						for (let c = 0; c < 2; c++) {
							if (c === col && r <= span)
								cells.push(
									cell(
										r === 1 ? lines(`${id}m`, n) : para(`${id}e${r}`),
										r === 1 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>',
									),
								);
							else cells.push(cell(para(`${id}r${r}c${c}`)));
						}
						rows += row(cells.join(''));
					}
					out[id] = {
						group: 'vmerge',
						body: para(`${id}top`) + tbl(rows, 2, border) + para(`${id}after`),
					};
				}
	for (const n of [1, 2]) {
		const id = `vmsp_n${n}`;
		const a = 'w:after="120"';
		const rows =
			row(
				cell(lines(`${id}m`, n, '120'), '<w:vMerge w:val="restart"/>') + cell(para(`${id}r1`, a)),
			) +
			row(cell(para(`${id}e2`, a), '<w:vMerge/>') + cell(para(`${id}r2`, a))) +
			row(cell(para(`${id}r3a`, a)) + cell(para(`${id}r3b`, a)));
		out[id] = {
			group: 'vmerge',
			body: para(`${id}top`) + tbl(rows, 2, 'single:4') + para(`${id}after`),
		};
	}
	for (const [k, border] of [
		['none', ''],
		['s4', 'single:4'],
		['s12', 'single:12'],
		['s24', 'single:24'],
		['dbl4', 'double:4'],
	] as const) {
		const id = `bd_${k}`;
		const rows = [1, 2, 3].map((r) => row(cell(para(`${id}r${r}`, 'w:after="120"')))).join('');
		out[id] = {
			group: 'borders',
			body: para(`${id}top`) + tbl(rows, 1, border) + para(`${id}after`),
		};
	}
	for (const rule of ['atLeast', 'exact'])
		for (const h of [800, 200]) {
			const id = `rh_${rule}_${h}`;
			const tr = `<w:trHeight w:val="${h}" w:hRule="${rule}"/>`;
			const rows = row(cell(para(`${id}r1`)), tr) + row(cell(para(`${id}r2`)), tr);
			out[id] = {
				group: 'rowh',
				body: para(`${id}top`) + tbl(rows, 1, 'single:4') + para(`${id}after`),
			};
		}
	return out;
}

/** A `.docx` for one document of the table probe (compatibility mode 15). */
export function tableDocx(name: string): Promise<Uint8Array> {
	return packProbe(tableBodies()[name]!.body, undefined, 15, false);
}
