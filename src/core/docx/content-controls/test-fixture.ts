// Templates with content controls for the tests: `sdt(properties, content)` builds a `w:sdt`.
import JSZip from 'jszip';
import { WORD_2012_NS } from '../content-control';
import { template } from '../mail-merge/test-fixture';
import type { DocumentModel } from '../model';
import { loadDocx } from '../parse';

export { paragraph, run } from '../mail-merge/test-fixture';

let next = 100;
/** A control with `tag` (and an id of its own) around `content`. */
export const sdt = (tag: string | undefined, content: string, extra = '') =>
	`<w:sdt><w:sdtPr><w:id w:val="${next++}"/>${tag === undefined ? '' : `<w:tag w:val="${tag}"/>`}${extra}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
export const section = `<w15:repeatingSection xmlns:w15="${WORD_2012_NS}"/>`;
export const item = `<w15:repeatingSectionItem xmlns:w15="${WORD_2012_NS}"/>`;
export const styled = (text: string, props: string) =>
	`<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
export const cell = (inner: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${inner}</w:tc>`;
export const table = (rows: string) =>
	`<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>${rows}</w:tbl>`;

export { template };

/** Saves `model` with the template's writer and loads the result: its model and document XML. */
export async function saveAndReload(
	loaded: Pick<Awaited<ReturnType<typeof template>>, 'save'>,
	model: DocumentModel,
): Promise<{ model: DocumentModel; xml: string; bytes: Uint8Array }> {
	const bytes = await loaded.save(model);
	const xml =
		(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
	return { model: (await loadDocx(bytes)).model, xml, bytes };
}
