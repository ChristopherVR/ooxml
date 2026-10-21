import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { expectParagraph, must } from './test-support/access';
import type { TextRun } from './model';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (inner: string, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}${inner}</w:r>`;
const text = (value: string, props = '') => run(`<w:t xml:space="preserve">${value}</w:t>`, props);

/** A text form field as Word writes it, its result runs marked `w:noProof`. */
const textField = (result: string) =>
	run(
		'<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar>',
	) +
	run('<w:instrText xml:space="preserve"> FORMTEXT </w:instrText>') +
	run('<w:fldChar w:fldCharType="separate"/>') +
	result +
	run('<w:fldChar w:fldCharType="end"/>');
/** Word's result for a text field with no text: five en spaces, one run each. */
const BLANK = Array.from({ length: 5 }, () => text('\u2002', '<w:noProof/>')).join('');

async function load(paragraphXml: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, paragraph: expectParagraph(loaded.model.blocks[0]) };
}

async function saved(loaded: Awaited<ReturnType<typeof loadDocx>>) {
	const bytes = await loaded.save();
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	return { xml, reopened: expectParagraph((await loadDocx(bytes)).model.blocks[0]) };
}

const resultRuns = (runs: readonly TextRun[]) =>
	runs.filter((item) => item.field?.instr.trim() === 'FORMTEXT');

describe('changing the runs of a paragraph whose runs have properties the model does not hold', () => {
	it('fills a blank text field Word wrote and keeps its result unproofed', async () => {
		const { loaded, paragraph } = await load(`<w:p>${text('Name: ')}${textField(BLANK)}</w:p>`);
		// The editor holds the five equal runs as one text, and what is typed takes their place.
		const [first, ...rest] = resultRuns(paragraph.runs);
		paragraph.runs = paragraph.runs
			.filter((item) => !rest.includes(item))
			.map((item) => (item === first ? { ...item, text: 'Ada' } : item));
		const { xml, reopened } = await saved(loaded);
		const result = resultRuns(reopened.runs);
		expect(result.map((item) => item.text)).toEqual(['Ada']);
		expect(must(result[0]).sourceRunPropertiesXml).toContain('noProof');
		expect(xml).toContain('<w:ffData><w:name w:val="Name"/>');
		expect(xml).not.toContain('\u2002');
		// The runs after the result keep their own properties: none moves onto the field's end.
		expect(xml).toContain('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
	});

	it('keeps the next field as it was when the runs of a blank before it are joined', async () => {
		const second =
			run(
				'<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Date"/><w:enabled/><w:calcOnExit w:val="0"/><w:entryMacro w:val="Check"/><w:textInput/></w:ffData></w:fldChar>',
				'<w:noProof/>',
			) +
			run('<w:instrText xml:space="preserve"> FORMTEXT </w:instrText>') +
			run('<w:fldChar w:fldCharType="separate"/>') +
			BLANK +
			run('<w:fldChar w:fldCharType="end"/>');
		const { loaded, paragraph } = await load(
			`<w:p>${text('Name: ')}${textField(BLANK)}${text(' Date: ')}${second}</w:p>`,
		);
		const [first, ...rest] = resultRuns(paragraph.runs).slice(0, 5);
		paragraph.runs = paragraph.runs
			.filter((item) => !rest.includes(item))
			.map((item) => (item === first ? { ...item, text: 'Ada' } : item));
		const { xml, reopened } = await saved(loaded);
		expect(resultRuns(reopened.runs).map((item) => item.text)).toEqual([
			'Ada',
			...Array.from({ length: 5 }, () => '\u2002'),
		]);
		// The second field's data and its runs' properties stay its own, though its runs moved.
		expect(xml).toMatch(
			/<w:ffData[^>]*><w:name w:val="Date"\/><w:enabled\/><w:calcOnExit w:val="0"\/><w:entryMacro w:val="Check"\/><w:textInput\/><\/w:ffData>/,
		);
		expect(xml.match(/<w:noProof\/>/g)).toHaveLength(1 + 1 + 5);
		expect(xml).toContain('<w:t xml:space="preserve"> Date: </w:t></w:r><w:r><w:rPr');
	});

	it('takes text typed at the end of a run with such properties into a run of its own', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${textField(text('Ada', '<w:noProof/>'))}${text(' after')}</w:p>`,
		);
		const index = paragraph.runs.findIndex((item) => item.text === 'Ada');
		const ada = must(paragraph.runs[index]);
		paragraph.runs.splice(index + 1, 0, { ...ada, text: ' Lovelace' });
		const { xml, reopened } = await saved(loaded);
		const result = resultRuns(reopened.runs);
		expect(result.map((item) => item.text).join('')).toBe('Ada Lovelace');
		for (const item of result) expect(item.sourceRunPropertiesXml).toContain('noProof');
		expect(xml).toContain('<w:t xml:space="preserve"> after</w:t></w:r></w:p>');
		expect(xml.match(/<w:noProof\/>/g)).toHaveLength(result.length);
	});

	it('refuses to join runs whose properties the model does not hold differ', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${text('Plain ')}${text('one', '<w:noProof/>')}${text(' two', '<w:outline/>')}</w:p>`,
		);
		const [plain, one] = paragraph.runs;
		paragraph.runs = [must(plain), { ...must(one), text: 'one two' }];
		await expect(loaded.save()).rejects.toThrow('changing run boundaries');
	});
});
