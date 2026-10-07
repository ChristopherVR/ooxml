import JSZip from 'jszip';
import { restartFixture } from './restart-fixture';

export async function paragraphBoundaryFixture(action: 'split' | 'join'): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await restartFixture(undefined));
	const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const p = (text: string, align: string) =>
		`<w:p><w:pPr><w:jc w:val="${align}"/><w:spacing w:after="120" data-extra="keep"/><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${action === 'split' ? p('Hello', 'center') : p('He', 'center') + p('llo', 'center')}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
