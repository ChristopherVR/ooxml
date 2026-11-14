// Test-only: small DOCX packages and body-text helpers for the document comparison tests.
import JSZip from 'jszip';
import type { Block, DocumentModel } from '../model';

const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** A package holding only `word/document.xml` with this body. */
export function docx(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** One plain paragraph per text. */
export const p = (...texts: string[]): string =>
	texts.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');

/** A plain table, one paragraph per cell. */
export const tbl = (...rows: string[][]): string =>
	`<w:tbl><w:tblPr/><w:tblGrid>${rows[0]!.map(() => '<w:gridCol w:w="2000"/>').join('')}</w:tblGrid>${rows
		.map((cells) => `<w:tr>${cells.map((cell) => `<w:tc>${p(cell)}</w:tc>`).join('')}</w:tr>`)
		.join('')}</w:tbl>`;

export const COMPARE_OPTIONS = { author: 'Compare Author', date: '2026-01-01T00:00:00Z' };

/** Paragraph texts of a story, tables flattened cell by cell. */
export function texts(blocks: Block[]): string[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block.runs.map((run) => run.text).join('')]
			: block.rows.flatMap((row) => row.flatMap((cell) => texts(cell.paragraphs))),
	);
}
export const bodyText = (model: DocumentModel): string[] => texts(model.blocks);
