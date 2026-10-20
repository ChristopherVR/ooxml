import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { at, expectParagraph } from './test-support/access';
import { symbolOnly } from './special-characters';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function load(paragraphXml: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, paragraph: expectParagraph(loaded.model.blocks[0]) };
}

async function savedParagraph(loaded: Awaited<ReturnType<typeof loadDocx>>): Promise<string> {
	const xml = await (
		await JSZip.loadAsync(await loaded.save())
	)
		.file('word/document.xml')!
		.async('string');
	return xml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ?? '';
}

/** Asserts that every needle occurs in `xml`, in this order. */
function expectInOrder(xml: string, needles: string[]) {
	let from = 0;
	for (const needle of needles) {
		const index = xml.indexOf(needle, from);
		expect(index, `${needle} after offset ${from} in ${xml}`).toBeGreaterThanOrEqual(0);
		from = index + needle.length;
	}
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

describe('editing paragraphs that hold position markers', () => {
	it('keeps spelling and grammar marks where they were', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A ')}<w:proofErr w:type="spellStart"/>${run('speling')}<w:proofErr w:type="spellEnd"/>` +
				`<w:proofErr w:type="gramStart"/>${run(' are here')}<w:proofErr w:type="gramEnd"/></w:p>`,
		);
		at(paragraph.runs, 0).text = 'The ';
		expectInOrder(await savedParagraph(loaded), [
			'>The <',
			'<w:proofErr w:type="spellStart"/>',
			'>speling<',
			'<w:proofErr w:type="spellEnd"/>',
			'<w:proofErr w:type="gramStart"/>',
			'> are here<',
			'<w:proofErr w:type="gramEnd"/>',
		]);
	});

	it('keeps an editing permission range around the text it covers', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:permStart w:id="4" w:edGrp="everyone"/>${run('editable')}<w:permEnd w:id="4"/>${run(' fixed')}</w:p>`,
		);
		at(paragraph.runs, 2).text = ' still fixed';
		expectInOrder(await savedParagraph(loaded), [
			'>Fixed <',
			'<w:permStart w:id="4" w:edGrp="everyone"/>',
			'>editable<',
			'<w:permEnd w:id="4"/>',
			'> still fixed<',
		]);
	});

	it('keeps a bookmark around the run it marks instead of widening it to the paragraph', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Label: ')}<w:bookmarkStart w:id="3" w:name="Value"/>${run('old')}<w:bookmarkEnd w:id="3"/></w:p>`,
		);
		expect(paragraph.bookmarks).toEqual(['Value']);
		at(paragraph.runs, 1).text = 'new';
		expectInOrder(await savedParagraph(loaded), [
			'>Label: <',
			'<w:bookmarkStart w:id="3" w:name="Value"/>',
			'>new<',
			'<w:bookmarkEnd w:id="3"/>',
		]);
	});

	it('brackets the whole paragraph with a bookmark added to the model', async () => {
		const { loaded, paragraph } = await load(`<w:p>${run('One')}${run(' two')}</w:p>`);
		paragraph.bookmarks = ['Added'];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['<w:bookmarkStart', 'w:name="Added"', '>One<', '> two<', '<w:bookmarkEnd']);
	});

	it('keeps marks at their characters when an editor merges the runs around them', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A ')}<w:proofErr w:type="spellStart"/>${run('speling')}<w:proofErr w:type="spellEnd"/>` +
				`<w:permStart w:id="4" w:edGrp="everyone"/>${run(' mistake')}<w:permEnd w:id="4"/></w:p>`,
		);
		paragraph.runs = [{ text: 'A speling misstake' }];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>A <',
			'<w:proofErr w:type="spellStart"/>',
			'>speling<',
			'<w:proofErr w:type="spellEnd"/>',
			'<w:permStart w:id="4" w:edGrp="everyone"/>',
			'> misstake<',
			'<w:permEnd w:id="4"/>',
		]);
		expect(xml.match(/<w:r>/g)).toHaveLength(3);
	});

	it('keeps text typed right after the end of a bookmark outside it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Meet in ')}<w:bookmarkStart w:id="2" w:name="Place"/>${run('Paris')}<w:bookmarkEnd w:id="2"/>${run(' soon')}</w:p>`,
		);
		paragraph.runs = [{ text: 'Meet in Paris! soon' }];
		expectInOrder(await savedParagraph(loaded), [
			'>Meet in <',
			'<w:bookmarkStart w:id="2" w:name="Place"/>',
			'>Paris<',
			'<w:bookmarkEnd w:id="2"/>',
			'>! soon<',
		]);
	});

	it('keeps a range that ends with the paragraph around text typed at its end', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:bookmarkStart w:id="2" w:name="Whole"/>${run('Paris')}<w:bookmarkEnd w:id="2"/></w:p>`,
		);
		paragraph.runs = [{ text: 'Paris!' }];
		expectInOrder(await savedParagraph(loaded), [
			'<w:bookmarkStart w:id="2" w:name="Whole"/>',
			'>Paris!<',
			'<w:bookmarkEnd w:id="2"/>',
		]);
	});

	it('keeps a bookmark around a field whose result is edited', async () => {
		const marker = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
		const { loaded, paragraph } = await load(
			`<w:p>${run('Name: ')}<w:bookmarkStart w:id="1" w:name="Name"/>${marker('begin')}` +
				`<w:r><w:instrText xml:space="preserve"> FILLIN "Name" </w:instrText></w:r>${marker('separate')}` +
				`${run('old')}${marker('end')}<w:bookmarkEnd w:id="1"/>${run(' done')}</w:p>`,
		);
		const result = paragraph.runs.findIndex((item) => item.text === 'old');
		at(paragraph.runs, result).text = 'Ada Lovelace';
		expectInOrder(await savedParagraph(loaded), [
			'>Name: <',
			'<w:bookmarkStart w:id="1" w:name="Name"/>',
			'w:fldCharType="begin"',
			'>Ada Lovelace<',
			'w:fldCharType="end"',
			'<w:bookmarkEnd w:id="1"/>',
			'> done<',
		]);
	});

	it('keeps an empty bookmark in order when text is typed exactly at it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:bookmarkStart w:id="0" w:name="Cursor"/><w:bookmarkEnd w:id="0"/>${run('B')}</w:p>`,
		);
		paragraph.runs = [{ text: 'A!B' }];
		const xml = await savedParagraph(loaded);
		expect(xml.indexOf('<w:bookmarkStart')).toBeLessThan(xml.indexOf('<w:bookmarkEnd'));
		expectInOrder(xml, [
			'>A',
			'<w:bookmarkStart w:id="0" w:name="Cursor"/>',
			'<w:bookmarkEnd w:id="0"/>',
		]);
	});

	it('never places a mark inside a surrogate pair', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('\u{1F600} ')}<w:bookmarkStart w:id="1" w:name="Word"/>${run('word')}<w:bookmarkEnd w:id="1"/></w:p>`,
		);
		paragraph.runs = [{ text: '\u{1F601} edit' }];
		const xml = await savedParagraph(loaded);
		expect(xml).toContain('\u{1F601}');
		expect(xml).not.toMatch(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/);
		expect(xml).not.toMatch(/(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/);
	});

	it('carries marks inside hyperlinks and content controls', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="target"><w:proofErr w:type="spellStart"/>${run('linkk')}<w:proofErr w:type="spellEnd"/></w:hyperlink>` +
				`<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="9" w:name="Inside"/>${run(' value')}<w:bookmarkEnd w:id="9"/></w:sdtContent></w:sdt>` +
				`${run(' after')}</w:p>`,
		);
		at(paragraph.runs, 2).text = ' changed';
		expectInOrder(await savedParagraph(loaded), [
			'<w:proofErr w:type="spellStart"/>',
			'>linkk<',
			'<w:proofErr w:type="spellEnd"/>',
			'<w:bookmarkStart w:id="9" w:name="Inside"/>',
			'> value<',
			'<w:bookmarkEnd w:id="9"/>',
			'> changed<',
		]);
	});
});

describe('editing paragraphs with smart tags, symbols, positional tabs and special hyphens', () => {
	it('reads the text inside smart tags and inline custom XML and keeps it through an edit', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Meet in ')}<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">${run('Paris')}</w:smartTag>` +
				`<w:customXml w:element="when">${run(' today')}</w:customXml></w:p>`,
		);
		expect(paragraph.runs.map((item) => item.text).join('')).toBe('Meet in Paris today');
		at(paragraph.runs, 0).text = 'Meet at ';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['>Meet at <', '>Paris<', '> today<']);
		expect(xml).not.toMatch(/w:smartTag|w:customXml/);
		expect(loaded.model.warnings.join(' ')).toContain('Smart tags and custom XML markup');
	});

	it('edits a paragraph whose custom XML and smart tag carry their own properties', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Ref ')}<w:customXml w:element="ref"><w:customXmlPr><w:attr w:name="kind" w:val="doc"/></w:customXmlPr>${run('A-1')}</w:customXml>` +
				`<w:smartTag w:element="place"><w:smartTagPr><w:attr w:name="x" w:val="y"/></w:smartTagPr>${run(' Paris')}</w:smartTag></w:p>`,
		);
		at(paragraph.runs, 0).text = 'See ';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['>See <', '>A-1<', '> Paris<']);
		expect(xml).not.toMatch(/customXmlPr|smartTagPr/);
	});

	it('draws only a run made of its symbol in the symbol font', () => {
		const symbol = { font: 'Wingdings', char: 'F0FE' };
		expect(symbolOnly({ text: '\uF0FE', symbol })).toBe(true);
		expect(symbolOnly({ text: '\uF0FE ok', symbol })).toBe(false);
		expect(symbolOnly({ text: 'x' })).toBe(false);
	});

	it('keeps a symbol character', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Done ')}<w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:sym w:font="Wingdings" w:char="F0FE"/></w:r></w:p>`,
		);
		expect(at(paragraph.runs, 1)).toMatchObject({
			text: '',
			symbol: { font: 'Wingdings', char: 'F0FE' },
		});
		at(paragraph.runs, 0).text = 'Checked ';
		expectInOrder(await savedParagraph(loaded), [
			'>Checked <',
			'<w:sym w:font="Wingdings" w:char="F0FE"/>',
		]);
	});

	it('writes plain text when a symbol is replaced', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:r><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>`,
		);
		at(paragraph.runs, 0).text = '*';
		const xml = await savedParagraph(loaded);
		expect(xml).toContain('<w:t>*</w:t>');
		expect(xml).not.toContain('w:sym');
	});

	it('keeps a positional tab', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Left')}<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>${run('Right')}</w:p>`,
		);
		expect(at(paragraph.runs, 1)).toMatchObject({
			text: '\t',
			positionalTab: { alignment: 'right', relativeTo: 'margin', leader: 'none' },
		});
		at(paragraph.runs, 0).text = 'Title';
		expectInOrder(await savedParagraph(loaded), [
			'>Title<',
			'<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>',
			'>Right<',
		]);
	});

	it('reads soft and non-breaking hyphens and writes them back as Word elements', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:r><w:t>co</w:t><w:softHyphen/><w:t>operate</w:t></w:r>${run(' with ')}` +
				`<w:r><w:t>well</w:t><w:noBreakHyphen/><w:t>known</w:t></w:r></w:p>`,
		);
		expect(paragraph.runs.map((item) => item.text)).toEqual([
			'co\u00ADoperate',
			' with ',
			'well\u2011known',
		]);
		at(paragraph.runs, 1).text = ' and ';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>co<',
			'<w:softHyphen/>',
			'>operate<',
			'> and <',
			'>well<',
			'<w:noBreakHyphen/>',
			'>known<',
		]);
		expect(xml).not.toMatch(/[\u00AD\u2011]/);
	});

	it('keeps a symbol that shares its run with text typed after it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r></w:p>`,
		);
		at(paragraph.runs, 0).text = '\uF0FE ok';
		const saved = await loaded.save();
		const xml = await savedParagraph(loaded);
		expect(xml).toContain(
			'<w:r><w:sym w:font="Wingdings" w:char="F0FE"/><w:t xml:space="preserve"> ok</w:t></w:r>',
		);
		const reread = expectParagraph((await loadDocx(saved)).model.blocks[0]);
		expect(at(reread.runs, 0)).toMatchObject({
			text: '\uF0FE ok',
			symbol: { font: 'Wingdings', char: 'F0FE' },
		});
	});

	it('still refuses a run that holds two different symbols', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:r><w:sym w:font="Wingdings" w:char="F0FE"/><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>`,
		);
		at(paragraph.runs, 0).text = 'C';
		await expect(loaded.save()).rejects.toThrow('cannot safely relocate');
	});
});
