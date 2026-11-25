import JSZip from 'jszip';
import { loadDocx } from '../parse';
import { REL_NS, WORD_NS } from '../xml';

export const marker = (kind: 'begin' | 'separate' | 'end') =>
	`<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
export const code = (text: string) =>
	`<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
export const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
export const simple = (instruction: string, cached = 'saved') =>
	`<w:fldSimple w:instr='${instruction}'>${run(cached)}</w:fldSimple>`;
export const complex = (instruction: string, cached?: string) =>
	`${marker('begin')}${instruction}${cached === undefined ? '' : marker('separate') + run(cached)}${marker('end')}`;
export const paragraph = (inner: string) => `<w:p>${inner}</w:p>`;

export async function template(body: string, parts: Record<string, string> = {}) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}" xmlns:r="${REL_NS}"><w:body>${body}</w:body></w:document>`,
	);
	for (const [path, xml] of Object.entries(parts)) zip.file(path, xml);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return { bytes, ...(await loadDocx(bytes)) };
}
