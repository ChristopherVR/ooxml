import JSZip from 'jszip';
import { restartFixture } from './restart-fixture';

export interface NumberingStartCase {
	name: string;
	abstractRestart: number | null;
	abstractStart: number | null;
	fullLevel: boolean;
	fullStart: number | null;
	startOverride: number | null;
}

/** Synthetic independent list-start inputs, without cached labels or user content. */
export async function numberingStartFixture(item: NumberingStartCase): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await restartFixture(item.abstractRestart ?? undefined));
	let source = await zip.file('word/numbering.xml')!.async('string');
	source = source.replace(
		'<w:lvl w:ilvl="2"><w:start w:val="1"/>',
		item.abstractStart === null
			? '<w:lvl w:ilvl="2">'
			: `<w:lvl w:ilvl="2"><w:start w:val="${item.abstractStart}"/>`,
	);
	const start = item.fullStart === null ? '' : `<w:start w:val="${item.fullStart}"/>`;
	const level = item.fullLevel
		? `<w:lvl w:ilvl="2">${start}<w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/></w:lvl>`
		: '';
	const override =
		item.startOverride === null ? '' : `<w:startOverride w:val="${item.startOverride}"/>`;
	zip.file(
		'word/numbering.xml',
		source.replace(
			'<w:abstractNumId w:val="0"/></w:num>',
			`<w:abstractNumId w:val="0"/>${level || override ? `<w:lvlOverride w:ilvl="2">${override}${level}</w:lvlOverride>` : ''}</w:num>`,
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}
