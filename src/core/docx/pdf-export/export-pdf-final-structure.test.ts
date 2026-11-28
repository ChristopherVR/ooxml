import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { finalReviewStructure } from '../review-final-structure';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer = createFakeMeasurer();
const stamp = 'w:author="Author" w:date="2026-01-01T00:00:00Z"';
const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

async function docx(body: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file('word/numbering.xml', numbering);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
	);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

/** The painted text, grouped into lines by baseline (top to bottom), with each line's y. */
function lines(bytes: Uint8Array): { y: number; text: string }[] {
	const raw = Buffer.from(bytes).toString('latin1');
	const found: { x: number; y: number; text: string }[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		if (
			!/\/Filter \/FlateDecode/.test(match[1]!) ||
			/\/Subtype \/(Image|Form)|\/Length1|\/N 3/.test(match[1]!)
		)
			continue;
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const content = inflateSync(bytes.subarray(start, start + length)).toString('latin1');
		for (const op of content.matchAll(
			/1 0 0 -1 ([\d.]+) ([\d.]+) Tm[^(]*\(((?:\\.|[^\\)])*)\) Tj/g,
		))
			found.push({ x: Number(op[1]), y: Number(op[2]), text: op[3]! });
	}
	const rows = new Map<number, { x: number; text: string }[]>();
	for (const item of found) {
		const key = Math.round(item.y * 2) / 2;
		rows.set(key, [...(rows.get(key) ?? []), item]);
	}
	return [...rows.entries()]
		.sort(([a], [b]) => a - b)
		.map(([y, items]) => ({
			y,
			text: items
				.sort((a, b) => a.x - b.x)
				.map((item) => item.text)
				.join('')
				.replace(/\s+/g, ' ')
				.trim(),
		}));
}
const p = (text: string, numbered = false) =>
	`<w:p>${numbered ? '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const deletedMark = (id: number, content: string, numbered = false) =>
	`<w:p><w:pPr>${numbered ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : ''}<w:rPr><w:del w:id="${id}" ${stamp}/></w:rPr></w:pPr>${content}</w:p>`;
const deletedRun = (id: number, text: string) =>
	`<w:del w:id="${id}" ${stamp}><w:r><w:delText>${text}</w:delText></w:r></w:del>`;

describe('exportDocxPdf without markup: deleted structure', () => {
	it('joins a paragraph whose mark is deleted to the next one', async () => {
		const model = await docx(
			p('First.') +
				deletedMark(1, deletedRun(2, 'gone text')) +
				p('Third.') +
				p('Item A', true) +
				deletedMark(3, deletedRun(4, 'Item B'), true) +
				p('Item C', true) +
				deletedMark(5, '<w:r><w:t xml:space="preserve">Joined, </w:t></w:r>') +
				p('so one line.'),
		);
		const final = finalReviewStructure(model);
		expect(final.blocks).toHaveLength(5);
		// The source model is not changed.
		expect(model.blocks).toHaveLength(8);
		const clean = await docx(
			p('First.') + p('Third.') + p('Item A', true) + p('Item C', true) + p('Joined, so one line.'),
		);
		const exported = lines((await exportDocxPdf(model, { measurer })).bytes);
		// No empty line or label for the deleted marks, and the joined text sits on one line.
		expect(exported.map((line) => line.text)).toEqual(
			lines((await exportDocxPdf(clean, { measurer })).bytes).map((line) => line.text),
		);
		expect(exported.map((line) => line.y)).toEqual(
			lines((await exportDocxPdf(clean, { measurer })).bytes).map((line) => line.y),
		);
		expect(exported.map((line) => line.text)).toContain('Joined, so one line.');
		expect(exported.map((line) => line.text).join('|')).not.toContain('3.');
		const marked = lines((await exportDocxPdf(model, { measurer, markup: 'all' })).bytes);
		expect(marked.map((line) => line.text).join('|')).toContain('gone');
	});

	it('removes rows deleted with tracking', async () => {
		const plain = (text: string) => `<w:tc>${p(text)}</w:tc>`;
		const dead = (id: number, text: string) =>
			`<w:tc>${deletedMark(id, deletedRun(id + 10, text))}</w:tc>`;
		const table = (rows: string) =>
			`<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>${p('After.')}`;
		const row = (a: string, b: string) => `<w:tr>${plain(a)}${plain(b)}</w:tr>`;
		const model = await docx(
			table(
				row('Row1A', 'Row1B') +
					`<w:tr><w:trPr><w:del w:id="1" ${stamp}/></w:trPr>${dead(2, 'DeadA')}${dead(3, 'DeadB')}</w:tr>` +
					row('Row3A', 'Row3B'),
			),
		);
		const final = finalReviewStructure(model).blocks[0]!;
		expect(final.type === 'table' && final.rows).toHaveLength(2);
		expect(final.type === 'table' && final.rowProperties).toHaveLength(2);
		const clean = await docx(table(row('Row1A', 'Row1B') + row('Row3A', 'Row3B')));
		const exported = lines((await exportDocxPdf(model, { measurer })).bytes);
		// The text after the table sits where it does when the row was never there.
		expect(exported).toEqual(lines((await exportDocxPdf(clean, { measurer })).bytes));
		expect(exported.at(-1)?.text).toBe('After.');
		expect(exported.map((line) => line.text).join('|')).not.toContain('Dead');
	});
});
