import JSZip from 'jszip';
import { restartFixture } from './restart-fixture';

/** Isolates restart semantics: abstract, full override and startOverride all start at one. */
export async function overrideRestartFixture(
	abstractRestart: number | undefined,
	overrideRestart: number,
): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await restartFixture(abstractRestart));
	const source = await zip.file('word/numbering.xml')!.async('string');
	const level = `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="${overrideRestart}"/><w:lvlText w:val="%3."/></w:lvl>`;
	zip.file(
		'word/numbering.xml',
		source.replace(
			'<w:abstractNumId w:val="0"/></w:num>',
			`<w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="2"><w:startOverride w:val="1"/>${level}</w:lvlOverride></w:num>`,
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}
