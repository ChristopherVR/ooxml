import JSZip from 'jszip';
import { restartFixture } from './restart-fixture';
import { WORD_NS } from '../xml';

export const fieldLockCases = [
	{ instr: 'SEQ Figure', cache: '42', locked: true },
	{ instr: 'SEQ Figure', cache: '9', locked: false },
	{ instr: 'REF Target', cache: 'Old title', locked: true },
	{ instr: 'REF Target', cache: 'Old title', locked: false },
	{ instr: 'PAGEREF Target', cache: '9', locked: true },
	{ instr: 'PAGEREF Target', cache: '9', locked: false },
] as const;

/** Native controls use identical instructions/caches as either simple or complex fields. */
export async function fieldLockFixture(kind: 'simple' | 'complex'): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await restartFixture(undefined));
	const fields = fieldLockCases.map(({ instr, cache, locked }) => {
		const flags = `w:fldLock="${locked}" w:dirty="true"`;
		const result = `<w:r><w:t>${cache}</w:t></w:r>`;
		const field =
			kind === 'simple'
				? `<w:fldSimple w:instr=" ${instr} " ${flags}>${result}</w:fldSimple>`
				: `<w:r><w:fldChar w:fldCharType="begin" ${flags}/></w:r><w:r><w:instrText> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
		return `<w:p>${field}</w:p>`;
	});
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>New title</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>${fields.join('')}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
