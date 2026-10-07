import JSZip from 'jszip';
import { restartFixture } from './restart-fixture';

export const hexStarts = [
	0, 1, 9, 10, 15, 16, 255, 256, 32767, 32768, 65535, 65536, 65537, 100000, 2147483647,
];
export const hexCases = [
	...hexStarts.map((value) => ({ value, template: '%1' })),
	{ value: 65535, template: '%1.' },
	{ value: 65536, template: '%1.' },
	{ value: 65535, template: '[%1]' },
	{ value: 65536, template: '[%1]' },
];

/** Independent lists isolate formatting from continuation/restart semantics. */
export async function hexNumberingFixture(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await restartFixture(undefined));
	const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${hexCases.map(({ value }, index) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${index + 1}"/></w:numPr></w:pPr><w:r><w:t>Value ${value}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<w:numbering xmlns:w="${w}">${hexCases.map(({ value, template }, index) => `<w:abstractNum w:abstractNumId="${index}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="${value}"/><w:numFmt w:val="hex"/><w:lvlText w:val="${template}"/></w:lvl></w:abstractNum>`).join('')}${hexCases.map((_, index) => `<w:num w:numId="${index + 1}"><w:abstractNumId w:val="${index}"/></w:num>`).join('')}</w:numbering>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
