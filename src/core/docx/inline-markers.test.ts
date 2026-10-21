import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { at, expectParagraph, lastOf } from './test-support/access';
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

/** The first `w:<tag>` element in `xml`, as text (none of the tags used here nest). */
const element = (xml: string, tag: string) =>
	xml.match(new RegExp(`<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`))?.[0] ?? '';

/** Whether some `w:<tag>` element in `xml` holds every needle. */
const heldBy = (xml: string, tag: string, needles: string[]) =>
	[...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`, 'g'))].some(([found]) =>
		needles.every((needle) => found.includes(needle)),
	);

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

	it('keeps a wrapper and a range around untouched text between two large edits', async () => {
		const [before, after] = ['a'.repeat(130), 'b'.repeat(130)];
		const { loaded, paragraph } = await load(
			`<w:p>${run(`${before} `)}<w:permStart w:id="4" w:edGrp="everyone"/>` +
				`<w:customXml w:uri="urn:x" w:element="name">${run('keep')}</w:customXml>` +
				`<w:permEnd w:id="4"/>${run(` ${after}`)}</w:p>`,
		);
		at(paragraph.runs, 0).text = `${'x'.repeat(130)} `;
		lastOf(paragraph.runs).text = ` ${'y'.repeat(130)}`;
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			`>${'x'.repeat(130)} <`,
			'<w:permStart w:id="4" w:edGrp="everyone"/>',
			'<w:customXml w:uri="urn:x" w:element="name">',
			'>keep<',
			'</w:customXml>',
			'<w:permEnd w:id="4"/>',
			`> ${'y'.repeat(130)}<`,
		]);
	});

	it('keeps markers inside the tracked change they were in', async () => {
		const revision = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins w:id="5" ${revision}><w:permStart w:id="4" w:edGrp="everyone"/>` +
				`${run('added')}<w:permEnd w:id="4"/></w:ins>` +
				`<w:del w:id="6" ${revision}><w:bookmarkStart w:id="2" w:name="Gone"/>` +
				`<w:r><w:delText>old</w:delText></w:r><w:bookmarkEnd w:id="2"/></w:del>` +
				`<w:ins w:id="7" ${revision}>${run('one')}<w:bookmarkStart w:id="3" w:name="Mid"/>` +
				`${run('two')}<w:bookmarkEnd w:id="3"/></w:ins>${run(' tail')}</w:p>`,
		);
		at(paragraph.runs, 0).text = 'Still fixed ';
		const xml = await savedParagraph(loaded);
		expect(xml).toContain('>Still fixed <');
		expect(heldBy(xml, 'ins', ['<w:permStart w:id="4"', '>added<', '<w:permEnd w:id="4"/>'])).toBe(
			true,
		);
		expect(
			heldBy(xml, 'del', ['<w:bookmarkStart w:id="2"', '>old<', '<w:bookmarkEnd w:id="2"/>']),
		).toBe(true);
		expect(heldBy(xml, 'ins', ['<w:bookmarkStart w:id="3"'])).toBe(true);
		expect(heldBy(xml, 'ins', ['<w:bookmarkEnd w:id="3"/>'])).toBe(true);
		expectInOrder(xml, [
			'>one<',
			'<w:bookmarkStart w:id="3"',
			'>two<',
			'<w:bookmarkEnd w:id="3"/>',
		]);
	});

	it('keeps an empty range inside a tracked change in order', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins ${revision}>${run('one')}` +
				`<w:bookmarkStart w:id="8" w:name="Empty"/><w:bookmarkEnd w:id="8"/>${run('two')}</w:ins></w:p>`,
		);
		at(paragraph.runs, 0).text = 'Still fixed ';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>one<',
			'<w:bookmarkStart w:id="8"',
			'<w:bookmarkEnd w:id="8"/>',
			'>two<',
		]);
		expect(heldBy(xml, 'ins', ['<w:bookmarkStart w:id="8"', '<w:bookmarkEnd w:id="8"/>'])).toBe(
			true,
		);
	});

	it('keeps ranges at their characters inside a tracked change whose runs the editor merges', async () => {
		const revision = 'w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins w:id="5" ${revision}>${run('left')}<w:bookmarkStart w:id="3" w:name="Mid"/>` +
				`${run('right')}<w:bookmarkEnd w:id="3"/><w:permStart w:id="4" w:edGrp="everyone"/>${run('more')}` +
				`<w:permEnd w:id="4"/></w:ins><w:del w:id="6" ${revision}><w:r><w:delText>was</w:delText></w:r>` +
				`<w:bookmarkStart w:id="2" w:name="Gone"/><w:r><w:delText>here</w:delText></w:r>` +
				`<w:bookmarkEnd w:id="2"/></w:del>${run(' tail')}</w:p>`,
		);
		// The editor holds the equally formatted runs of one change as one text.
		const [fixed, left, , , was, , tail] = paragraph.runs;
		paragraph.runs = [
			{ ...fixed!, text: 'Still fixed ' },
			{ ...left!, text: 'leftrightmore' },
			{ ...was!, text: 'washere' },
			tail!,
		];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>Still fixed <',
			'>left<',
			'<w:bookmarkStart w:id="3" w:name="Mid"/>',
			'>right<',
			'<w:bookmarkEnd w:id="3"/>',
			'<w:permStart w:id="4" w:edGrp="everyone"/>',
			'>more<',
			'<w:permEnd w:id="4"/>',
			'<w:delText>was</w:delText>',
			'<w:bookmarkStart w:id="2" w:name="Gone"/>',
			'<w:delText>here</w:delText>',
			'<w:bookmarkEnd w:id="2"/>',
			'> tail<',
		]);
		expect(heldBy(xml, 'ins', ['>left<', '<w:bookmarkStart w:id="3"', '>right<'])).toBe(true);
		expect(heldBy(xml, 'del', ['>was<', '<w:bookmarkStart w:id="2"', '>here<'])).toBe(true);
	});

	it('keeps a smart tag around its text inside a tracked change whose runs the editor merges', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins ${revision}>${run('Meet in ')}` +
				`<w:smartTag w:element="City">${run('Paris')}</w:smartTag>${run(' today')}</w:ins></w:p>`,
		);
		const [fixed, meet] = paragraph.runs;
		paragraph.runs = [
			{ ...fixed!, text: 'Still fixed ' },
			{ ...meet!, text: 'Meet in Paris today' },
		];
		const xml = await savedParagraph(loaded);
		expect(heldBy(xml, 'ins', ['>Meet in <', '<w:smartTag w:element="City">', '>Paris<'])).toBe(
			true,
		);
		expectInOrder(xml, ['>Meet in <', '<w:smartTag w:element="City">', '>Paris<', '</w:smartTag>']);
		expect(element(xml, 'smartTag')).not.toContain('today');
	});

	it('keeps smart tags that start and end a tracked change whose runs the editor merges', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins ${revision}><w:smartTag w:element="Day">${run('Today')}</w:smartTag>` +
				`${run(' in ')}<w:smartTag w:element="City">${run('Paris')}</w:smartTag></w:ins>${run(' tail')}</w:p>`,
		);
		const [fixed, today, , , tail] = paragraph.runs;
		paragraph.runs = [
			{ ...fixed!, text: 'Still fixed ' },
			{ ...today!, text: 'Today in Paris' },
			tail!,
		];
		const xml = await savedParagraph(loaded);
		expect(
			heldBy(xml, 'ins', [
				'<w:smartTag w:element="Day">',
				'>Today<',
				'<w:smartTag w:element="City">',
			]),
		).toBe(true);
		expectInOrder(xml, [
			'<w:smartTag w:element="Day">',
			'>Today<',
			'</w:smartTag>',
			'> in <',
			'<w:smartTag w:element="City">',
			'>Paris<',
			'</w:smartTag>',
			'</w:ins>',
			'> tail<',
		]);
	});

	it('keeps a smart tag around the tracked change it wraps', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:smartTag w:element="City"><w:ins ${revision}>${run('Paris')}</w:ins></w:smartTag></w:p>`,
		);
		at(paragraph.runs, 0).text = 'Still fixed ';
		expectInOrder(await savedParagraph(loaded), [
			'>Still fixed <',
			'<w:smartTag w:element="City">',
			'<w:ins ',
			'>Paris<',
			'</w:ins>',
			'</w:smartTag>',
		]);
	});

	it('keeps a smart tag around two differently formatted runs of one tracked change', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, paragraph } = await load(
			`<w:p>${run('Fixed ')}<w:ins ${revision}><w:smartTag w:element="City">${run('Pa')}` +
				`<w:r><w:rPr><w:b/></w:rPr><w:t>ris</w:t></w:r></w:smartTag></w:ins></w:p>`,
		);
		at(paragraph.runs, 0).text = 'Still fixed ';
		expectInOrder(await savedParagraph(loaded), [
			'>Still fixed <',
			'<w:smartTag w:element="City">',
			'>Pa<',
			'>ris<',
			'</w:smartTag>',
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

	it('keeps a bookmark in place when the paragraph is edited on both sides of it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Xa ')}<w:bookmarkStart w:id="3" w:name="Mid"/>${run('b')}<w:bookmarkEnd w:id="3"/>${run(' cY')}</w:p>`,
		);
		at(paragraph.runs, 0).text = 'Za ';
		at(paragraph.runs, 2).text = ' cW';
		expectInOrder(await savedParagraph(loaded), [
			'>Za <',
			'<w:bookmarkStart w:id="3" w:name="Mid"/>',
			'>b<',
			'<w:bookmarkEnd w:id="3"/>',
			'> cW<',
		]);
	});

	it('keeps an editing permission to its text when the words around it change', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('left ')}<w:permStart w:id="1" w:edGrp="everyone"/>${run('value')}<w:permEnd w:id="1"/>${run(' right')}</w:p>`,
		);
		paragraph.runs = [{ text: 'LEFT value RIGHT' }];
		expectInOrder(await savedParagraph(loaded), [
			'>LEFT <',
			'<w:permStart w:id="1" w:edGrp="everyone"/>',
			'>value<',
			'<w:permEnd w:id="1"/>',
			'> RIGHT<',
		]);
	});

	it('keeps a character typed before a range outside it when it repeats the next one', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('ab')}<w:bookmarkStart w:id="1" w:name="Tail"/>${run('c')}<w:bookmarkEnd w:id="1"/></w:p>`,
		);
		paragraph.runs = [{ text: 'abcc' }];
		expectInOrder(await savedParagraph(loaded), [
			'>abc<',
			'<w:bookmarkStart w:id="1" w:name="Tail"/>',
			'>c<',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});

	it('keeps a repeated character typed after a range at the paragraph start outside it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:bookmarkStart w:id="1" w:name="Lead"/>${run('a')}<w:bookmarkEnd w:id="1"/>${run('aa')}</w:p>`,
		);
		paragraph.runs = [{ text: 'aaaa' }];
		expectInOrder(await savedParagraph(loaded), [
			'<w:bookmarkStart w:id="1" w:name="Lead"/>',
			'>a<',
			'<w:bookmarkEnd w:id="1"/>',
			'>aaa<',
		]);
	});

	it('takes a deleted repeated character from outside a range rather than from it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('c')}<w:bookmarkStart w:id="1" w:name="Pair"/>${run('cc')}<w:bookmarkEnd w:id="1"/>${run('a')}</w:p>`,
		);
		paragraph.runs = [{ text: 'cca' }];
		expectInOrder(await savedParagraph(loaded), [
			'<w:bookmarkStart w:id="1" w:name="Pair"/>',
			'>cc<',
			'<w:bookmarkEnd w:id="1"/>',
			'>a<',
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
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'<w:proofErr w:type="spellStart"/>',
			'>linkk<',
			'<w:proofErr w:type="spellEnd"/>',
			'<w:bookmarkStart w:id="9" w:name="Inside"/>',
			'> value<',
			'<w:bookmarkEnd w:id="9"/>',
			'> changed<',
		]);
		expect(element(xml, 'hyperlink')).toContain('<w:proofErr w:type="spellEnd"/>');
		expect(element(xml, 'sdtContent')).toContain('<w:bookmarkEnd w:id="9"/>');
		expect(element(xml, 'sdtContent')).not.toContain('proofErr');
	});

	it('keeps a mark that closes a hyperlink at the end of the paragraph inside it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="target"><w:proofErr w:type="spellStart"/>${run('linkk')}<w:proofErr w:type="spellEnd"/></w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).text = 'link';
		const xml = await savedParagraph(loaded);
		expect(element(xml, 'hyperlink')).toContain('<w:proofErr w:type="spellEnd"/>');
		expect(xml.slice(xml.indexOf('</w:hyperlink>'))).not.toContain('proofErr');
	});

	it('keeps marks inside a hyperlink whose target changed', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="target"><w:proofErr w:type="spellStart"/>${run('linkk')}<w:proofErr w:type="spellEnd"/></w:hyperlink>${run(' after')}</w:p>`,
		);
		at(paragraph.runs, 0).link = { anchor: 'other' };
		const link = element(await savedParagraph(loaded), 'hyperlink');
		expect(link).toContain('w:anchor="other"');
		expectInOrder(link, [
			'<w:proofErr w:type="spellStart"/>',
			'>linkk<',
			'<w:proofErr w:type="spellEnd"/>',
		]);
	});

	it('keeps marks at the end of a container in it, not in the container after it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:bookmarkStart w:id="2" w:name="Value"/><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${run('value')}<w:bookmarkEnd w:id="2"/></w:sdtContent></w:sdt>` +
				`<w:hyperlink w:anchor="target">${run(' link')}<w:bookmarkStart w:id="3" w:name="Next"/></w:hyperlink>` +
				`<w:sdt><w:sdtPr><w:id w:val="6"/></w:sdtPr><w:sdtContent>${run(' more')}<w:bookmarkEnd w:id="3"/></w:sdtContent></w:sdt>${run('!')}</w:p>`,
		);
		lastOf(paragraph.runs).text = '?';
		const xml = await savedParagraph(loaded);
		expect(element(xml, 'sdtContent')).toContain('<w:bookmarkEnd w:id="2"/>');
		expect(element(xml, 'hyperlink')).toContain('<w:bookmarkStart w:id="3" w:name="Next"/>');
		expect(element(xml, 'hyperlink')).not.toContain('bookmarkEnd');
	});

	it('keeps a hyperlink and the marks in it when an editor merges the runs before it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Label')}${run(': ')}<w:hyperlink w:anchor="target" w:tgtFrame="_blank">` +
				`<w:bookmarkStart w:id="4" w:name="Link"/>${run('link')}<w:bookmarkEnd w:id="4"/></w:hyperlink></w:p>`,
		);
		paragraph.runs = [{ text: 'Label - ' }, ...paragraph.runs.filter((item) => item.link)];
		const link = element(await savedParagraph(loaded), 'hyperlink');
		expect(link).toContain('w:tgtFrame="_blank"');
		expect(link).not.toContain('w:history');
		expectInOrder(link, [
			'<w:bookmarkStart w:id="4" w:name="Link"/>',
			'>link<',
			'<w:bookmarkEnd w:id="4"/>',
		]);
	});

	it('closes a range at the end of a hyperlink run that an earlier mark split', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t"><w:bookmarkStart w:id="1" w:name="All"/>${run('ab')}` +
				`<w:proofErr w:type="spellStart"/>${run('cd')}<w:bookmarkEnd w:id="1"/></w:hyperlink></w:p>`,
		);
		paragraph.runs = [{ ...at(paragraph.runs, 0), text: 'abcd' }];
		expectInOrder(element(await savedParagraph(loaded), 'hyperlink'), [
			'<w:bookmarkStart w:id="1" w:name="All"/>',
			'>ab<',
			'<w:proofErr w:type="spellStart"/>',
			'>cd<',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});

	it('reuses the hyperlink that keeps its text when another link to the same place is removed', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t" w:tgtFrame="_blank">${run('one')}</w:hyperlink>${run(' and ')}` +
				`<w:hyperlink w:anchor="t"><w:bookmarkStart w:id="2" w:name="Two"/>${run('two')}<w:bookmarkEnd w:id="2"/></w:hyperlink></w:p>`,
		);
		delete at(paragraph.runs, 0).link;
		const xml = await savedParagraph(loaded);
		expect(xml.match(/<w:hyperlink\b/g)).toHaveLength(1);
		const link = element(xml, 'hyperlink');
		expect(link).not.toContain('tgtFrame');
		expectInOrder(link, [
			'<w:bookmarkStart w:id="2" w:name="Two"/>',
			'>two<',
			'<w:bookmarkEnd w:id="2"/>',
		]);
	});

	it('gives a new link to the same place none of an existing hyperlink markup', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('word')}${run(' gap ')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank">` +
				`<w:bookmarkStart w:id="5" w:name="Link"/>${run('link')}<w:bookmarkEnd w:id="5"/></w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).link = { anchor: 't' };
		const xml = await savedParagraph(loaded);
		const links = xml.match(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).not.toContain('tgtFrame');
		expect(at(links, 1)).toContain('w:tgtFrame="_blank"');
		expectInOrder(at(links, 1), [
			'<w:bookmarkStart w:id="5" w:name="Link"/>',
			'>link<',
			'<w:bookmarkEnd w:id="5"/>',
		]);
	});

	it('keeps the markup of the link that stays when its neighbour to the same place is unlinked', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t" w:tgtFrame="_blank">${run('one ')}</w:hyperlink>` +
				`<w:hyperlink w:anchor="t">${run('two')}</w:hyperlink></w:p>`,
		);
		delete at(paragraph.runs, 1).link;
		const xml = await savedParagraph(loaded);
		expect(xml.match(/<w:hyperlink\b/g)).toHaveLength(1);
		expectInOrder(xml, [
			'<w:hyperlink w:anchor="t" w:tgtFrame="_blank">',
			'>one <',
			'</w:hyperlink>',
			'>two<',
		]);
	});

	it('keeps a bookmark on its text when a link moves to an equal label elsewhere', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('link')}${run(' gap ')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank">` +
				`<w:bookmarkStart w:id="2" w:name="B"/>${run('link')}<w:bookmarkEnd w:id="2"/></w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).link = { anchor: 't' };
		delete at(paragraph.runs, 2).link;
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'<w:hyperlink w:anchor="t"',
			'>link<',
			'</w:hyperlink>',
			'> gap <',
			'<w:bookmarkStart w:id="2" w:name="B"/>',
			'>link<',
			'<w:bookmarkEnd w:id="2"/>',
		]);
		expect(xml).not.toContain('tgtFrame');
	});

	it('gives a link added to surviving text none of the markup of a deleted link', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('X ')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank">${run('one')}</w:hyperlink>${run(' two')}</w:p>`,
		);
		paragraph.runs = [{ ...at(paragraph.runs, 2), text: 'two', link: { anchor: 't' } }];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['<w:hyperlink w:anchor="t"', '>two<', '</w:hyperlink>']);
		expect(xml).not.toContain('tgtFrame');
	});

	it('keeps the markup of a link whose text survives when a new label replaces the text before it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('X')}${run('Y')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank">${run('one')}</w:hyperlink>${run(' gap ')}</w:p>`,
		);
		paragraph.runs = [
			{ text: 'NEW', link: { anchor: 't' } },
			{ ...at(paragraph.runs, 3) },
			{ ...at(paragraph.runs, 2), text: 'one' },
		];
		const links =
			(await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).not.toContain('tgtFrame');
		expect(at(links, 1)).toContain('w:tgtFrame="_blank"');
		expect(at(links, 1)).toContain('>one<');
	});

	it('keeps the markup of a hyperlink whose whole label is retyped', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('See ')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank" w:docLocation="here">` +
				`<w:bookmarkStart w:id="1" w:name="Label"/>${run('Docs')}<w:bookmarkEnd w:id="1"/></w:hyperlink>${run(' now')}</w:p>`,
		);
		at(paragraph.runs, 1).text = 'Help';
		const xml = await savedParagraph(loaded);
		expect(xml.match(/<w:hyperlink\b/g)).toHaveLength(1);
		const link = element(xml, 'hyperlink');
		expect(link).toContain('w:tgtFrame="_blank"');
		expect(link).toContain('w:docLocation="here"');
		expect(link).not.toContain('w:history');
		expectInOrder(link, [
			'<w:bookmarkStart w:id="1" w:name="Label"/>',
			'>Help<',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});

	it('gives a retyped label the markup of its own link, not of a link to another place', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="a" w:tgtFrame="_blank">${run('Alpha')}</w:hyperlink>` +
				`<w:hyperlink w:anchor="b" w:docLocation="here">${run('Beta')}</w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).text = 'One';
		at(paragraph.runs, 1).text = 'Two';
		const links = (await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).toContain('w:tgtFrame="_blank"');
		expect(at(links, 0)).not.toContain('docLocation');
		expect(at(links, 1)).toContain('w:docLocation="here"');
		expect(at(links, 1)).not.toContain('tgtFrame');
	});

	it('keeps each hyperlink its own markup when an edit swaps two of them', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="a" w:tgtFrame="_blank">${run('Alpha')}</w:hyperlink>${run(' and ')}` +
				`<w:hyperlink w:anchor="b" w:docLocation="here">${run('Beta')}</w:hyperlink></w:p>`,
		);
		const [alpha, and, beta] = [
			at(paragraph.runs, 0),
			at(paragraph.runs, 1),
			at(paragraph.runs, 2),
		];
		paragraph.runs = [beta, and, alpha];
		const links = (await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).toContain('w:docLocation="here"');
		expect(at(links, 1)).toContain('w:tgtFrame="_blank"');
	});

	it('gives a retargeted hyperlink markup to no new link to its old target', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('word')}${run(' gap ')}<w:hyperlink w:anchor="a" w:tgtFrame="_blank">` +
				`<w:bookmarkStart w:id="6" w:name="Link"/>${run('link')}<w:bookmarkEnd w:id="6"/></w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).link = { anchor: 'a' };
		at(paragraph.runs, 2).link = { anchor: 'b' };
		const links =
			(await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).not.toContain('tgtFrame');
		expect(at(links, 1)).toContain('w:anchor="b"');
		expectInOrder(at(links, 1), [
			'<w:bookmarkStart w:id="6" w:name="Link"/>',
			'>link<',
			'<w:bookmarkEnd w:id="6"/>',
		]);
	});

	it('keeps a hyperlink markup on both parts when its middle is unlinked', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t" w:tgtFrame="_blank">${run('click')}${run(' here ')}` +
				`<w:bookmarkStart w:id="7" w:name="Now"/>${run('now')}<w:bookmarkEnd w:id="7"/></w:hyperlink></w:p>`,
		);
		delete at(paragraph.runs, 1).link;
		const links =
			(await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/g) ?? [];
		expect(links).toHaveLength(2);
		for (const link of links) expect(link).toContain('w:tgtFrame="_blank"');
		expectInOrder(at(links, 1), [
			'<w:bookmarkStart w:id="7" w:name="Now"/>',
			'>now<',
			'<w:bookmarkEnd w:id="7"/>',
		]);
	});

	it('keeps a retargeted hyperlink in place when a same label elsewhere takes its old target', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('link')}${run(' gap ')}<w:hyperlink w:anchor="a" w:tgtFrame="_blank">${run('link')}</w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).link = { anchor: 'a' };
		at(paragraph.runs, 2).link = { anchor: 'b' };
		const links = (await savedParagraph(loaded)).match(/<w:hyperlink\b[^>]*>/g) ?? [];
		expect(links).toHaveLength(2);
		expect(at(links, 0)).toContain('w:anchor="a"');
		expect(at(links, 0)).not.toContain('tgtFrame');
	});

	it('keeps a smart tag around the middle of a hyperlink that was unlinked', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t">${run('left ')}<w:smartTag w:element="x">${run('middle')}</w:smartTag>${run(' right')}</w:hyperlink></w:p>`,
		);
		delete at(paragraph.runs, 1).link;
		expectInOrder(await savedParagraph(loaded), [
			'</w:hyperlink>',
			'<w:smartTag w:element="x">',
			'>middle<',
			'</w:smartTag>',
			'<w:hyperlink',
		]);
	});

	it('keeps each content control its own marks when two with the same text swap places', async () => {
		const control = (id: number, name: string) =>
			`<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="${id}" w:name="${name}"/>${run('x')}<w:bookmarkEnd w:id="${id}"/></w:sdtContent></w:sdt>`;
		const { loaded, paragraph } = await load(
			`<w:p>${control(1, 'One')}${run(' ')}${control(2, 'Two')}</w:p>`,
		);
		paragraph.runs = [at(paragraph.runs, 2), at(paragraph.runs, 1), at(paragraph.runs, 0)];
		const controls = (await savedParagraph(loaded)).match(/<w:sdt>[\s\S]*?<\/w:sdt>/g) ?? [];
		expect(controls).toHaveLength(2);
		expect(at(controls, 0)).toContain('<w:id w:val="2"/>');
		expect(at(controls, 0)).toContain('w:name="Two"');
		expect(at(controls, 1)).toContain('w:name="One"');
	});

	it('keeps a bookmark around a content control that moves past a neighbour with the same letters', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="1" w:name="All"/>` +
				`${run('aaa')}<w:bookmarkEnd w:id="1"/></w:sdtContent></w:sdt>${run('a')}</w:p>`,
		);
		paragraph.runs = [at(paragraph.runs, 1), at(paragraph.runs, 0)];
		expectInOrder(element(await savedParagraph(loaded), 'sdt'), [
			'<w:bookmarkStart w:id="1" w:name="All"/>',
			'>aaa<',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});

	it('keeps a bookmark on its letters inside a content control that moved and changed', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent>${run('A')}<w:bookmarkStart w:id="4" w:name="Part"/>` +
				`${run('da')}<w:bookmarkEnd w:id="4"/></w:sdtContent></w:sdt>${run(' signed')}</w:p>`,
		);
		const [, , signed] = [at(paragraph.runs, 0), at(paragraph.runs, 1), at(paragraph.runs, 2)];
		const control = at(paragraph.runs, 0).contentControl!;
		paragraph.runs = [
			{ ...signed, text: ' signed by ' },
			{ text: 'Dr Ada', contentControl: control },
		];
		expectInOrder(element(await savedParagraph(loaded), 'sdt'), [
			'>Dr A<',
			'<w:bookmarkStart w:id="4" w:name="Part"/>',
			'>da<',
			'<w:bookmarkEnd w:id="4"/>',
		]);
	});

	it('keeps a bookmark inside a content control that moved past one character', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="5" w:name="Who"/>` +
				`${run('Ada')}<w:bookmarkEnd w:id="5"/></w:sdtContent></w:sdt>${run('!')}</w:p>`,
		);
		paragraph.runs = [at(paragraph.runs, 1), at(paragraph.runs, 0)];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>!<',
			'<w:sdt>',
			'<w:bookmarkStart w:id="5" w:name="Who"/>',
			'>Ada<',
			'<w:bookmarkEnd w:id="5"/>',
		]);
	});

	it('keeps text typed at the start of a content control outside the bookmark in it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('X')}<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="1" w:name="Who"/>` +
				`${run('Ada')}<w:bookmarkEnd w:id="1"/></w:sdtContent></w:sdt>${run('Y')}</w:p>`,
		);
		at(paragraph.runs, 1).text = 'Dr Ada';
		expectInOrder(element(await savedParagraph(loaded), 'sdt'), [
			'>Dr <',
			'<w:bookmarkStart w:id="1" w:name="Who"/>',
			'>Ada<',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});

	it('keeps a bookmark on one letter of a content control when a repeated letter is added', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('X')}<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>${run('a')}` +
				`<w:bookmarkStart w:id="2" w:name="Mid"/>${run('a')}<w:bookmarkEnd w:id="2"/>${run('a')}</w:sdtContent></w:sdt>${run('Y')}</w:p>`,
		);
		const control = at(paragraph.runs, 1).contentControl!;
		paragraph.runs = [
			at(paragraph.runs, 0),
			{ text: 'aaaa', contentControl: control },
			lastOf(paragraph.runs),
		];
		expect(element(await savedParagraph(loaded), 'sdt')).toMatch(
			/<w:bookmarkStart w:id="2" w:name="Mid"\/><w:r><w:t>a<\/w:t><\/w:r><w:bookmarkEnd w:id="2"\/>/,
		);
	});

	it('keeps a bookmark on the first letter of a content control at the paragraph start', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="2" w:name="Lead"/>` +
				`${run('a')}<w:bookmarkEnd w:id="2"/>${run('aa')}</w:sdtContent></w:sdt>${run('Y')}</w:p>`,
		);
		const control = at(paragraph.runs, 0).contentControl!;
		paragraph.runs = [{ text: 'aaaa', contentControl: control }, lastOf(paragraph.runs)];
		expect(element(await savedParagraph(loaded), 'sdt')).toMatch(
			/<w:bookmarkStart w:id="2" w:name="Lead"\/><w:r><w:t>a<\/w:t><\/w:r><w:bookmarkEnd w:id="2"\/>/,
		);
	});

	it('keeps the marks inside a content control that moved within the paragraph', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="4" w:name="Signer"/>` +
				`${run('Ada')}<w:bookmarkEnd w:id="4"/></w:sdtContent></w:sdt>${run(' signed')}</w:p>`,
		);
		const [name, signed] = [at(paragraph.runs, 0), at(paragraph.runs, 1)];
		paragraph.runs = [{ ...signed, text: 'Signed by ' }, name];
		expectInOrder(element(await savedParagraph(loaded), 'sdt'), [
			'<w:bookmarkStart w:id="4" w:name="Signer"/>',
			'>Ada<',
			'<w:bookmarkEnd w:id="4"/>',
		]);
	});

	it('keeps an empty content control its markup when text is typed into it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Name: ')}<w:sdt><w:sdtPr><w:id w:val="3"/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent/></w:sdt>${run(' end')}</w:p>`,
		);
		const value = paragraph.runs.findIndex((item) => item.contentControl);
		at(paragraph.runs, value).text = 'Ada';
		const control = element(await savedParagraph(loaded), 'sdt');
		expect(control).toContain('<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>');
		expect(control).toContain('>Ada<');
	});

	it('keeps a content control its markup when its old text ends up outside it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>` +
				`<w:sdtContent>${run('B')}</w:sdtContent></w:sdt></w:p>`,
		);
		at(paragraph.runs, 0).text = 'B';
		at(paragraph.runs, 1).text = 'A';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['>B<', '<w:sdt>', '<w:id w:val="7"/>', '<w:sdtEndPr>', '>A<', '</w:sdt>']);
		expect(element(xml, 'sdt')).toContain('<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>');
	});

	it('keeps a simple field its markup when its old result ends up outside it', async () => {
		const open = '<w:fldSimple w:instr=" TITLE " xmlns:x="urn:x" x:keep="1">';
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}${open}${run('B')}</w:fldSimple></w:p>`,
		);
		at(paragraph.runs, 0).text = 'B';
		at(paragraph.runs, 1).text = 'A';
		expectInOrder(await savedParagraph(loaded), ['>B<', open, '>A<', '</w:fldSimple>']);
	});

	it('keeps a bookmark on its text when a link moves from that text to other text', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t" w:tgtFrame="_blank"><w:bookmarkStart w:id="1" w:name="Word"/>` +
				`${run('one')}<w:bookmarkEnd w:id="1"/></w:hyperlink>${run(' gap ')}${run('two')}</w:p>`,
		);
		delete at(paragraph.runs, 0).link;
		at(paragraph.runs, 2).link = { anchor: 't' };
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'<w:bookmarkStart w:id="1" w:name="Word"/>',
			'>one<',
			'<w:bookmarkEnd w:id="1"/>',
			'> gap <',
			'<w:hyperlink w:anchor="t"',
			'>two<',
		]);
		expect(xml).not.toContain('tgtFrame');
	});

	it('keeps the order of marks that close at the same place in a hyperlink', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="target"><w:bookmarkStart w:id="1" w:name="A"/><w:bookmarkStart w:id="2" w:name="B"/>${run('text')}` +
				`<w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="1"/></w:hyperlink>${run(' tail')}</w:p>`,
		);
		at(paragraph.runs, 1).text = ' end';
		expectInOrder(element(await savedParagraph(loaded), 'hyperlink'), [
			'<w:bookmarkStart w:id="1" w:name="A"/>',
			'<w:bookmarkStart w:id="2" w:name="B"/>',
			'>text<',
			'<w:bookmarkEnd w:id="2"/>',
			'<w:bookmarkEnd w:id="1"/>',
		]);
	});
});

describe('editing paragraphs with smart tags, symbols, positional tabs and special hyphens', () => {
	it('reads the text of smart tags and inline custom XML and keeps them around text left as it was', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Meet in ')}<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">${run('Paris')}</w:smartTag>` +
				`<w:customXml w:element="when">${run(' today')}</w:customXml></w:p>`,
		);
		expect(paragraph.runs.map((item) => item.text).join('')).toBe('Meet in Paris today');
		at(paragraph.runs, 0).text = 'Meet at ';
		expectInOrder(await savedParagraph(loaded), [
			'>Meet at <',
			'<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">',
			'>Paris<',
			'</w:smartTag>',
			'<w:customXml w:element="when">',
			'> today<',
			'</w:customXml>',
		]);
		expect(loaded.model.warnings.join(' ')).toContain('Smart tags and custom XML markup');
	});

	it('keeps the properties of custom XML and smart tags around text left as it was', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Ref ')}<w:customXml w:element="ref"><w:customXmlPr><w:attr w:name="kind" w:val="doc"/></w:customXmlPr>${run('A-1')}</w:customXml>` +
				`<w:smartTag w:element="place"><w:smartTagPr><w:attr w:name="x" w:val="y"/></w:smartTagPr>${run(' Paris')}</w:smartTag></w:p>`,
		);
		at(paragraph.runs, 0).text = 'See ';
		expectInOrder(await savedParagraph(loaded), [
			'>See <',
			'<w:customXml w:element="ref"><w:customXmlPr><w:attr w:name="kind" w:val="doc"/></w:customXmlPr>',
			'>A-1<',
			'</w:customXml>',
			'<w:smartTag w:element="place"><w:smartTagPr><w:attr w:name="x" w:val="y"/></w:smartTagPr>',
			'> Paris<',
			'</w:smartTag>',
		]);
	});

	it('drops the markup of a smart tag whose text was edited, and keeps the text', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Meet in ')}<w:smartTag w:element="City">${run('Paris')}</w:smartTag>${run(' today')}</w:p>`,
		);
		at(paragraph.runs, 1).text = 'Lyon';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['>Meet in <', '>Lyon<', '> today<']);
		expect(xml).not.toContain('smartTag');
	});

	it('keeps a smart tag around its text when the editor merges its runs with their neighbours', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Meet in ')}<w:smartTag w:element="City">${run('Paris')}</w:smartTag>${run(' today')}</w:p>`,
		);
		paragraph.runs = [{ text: 'Meet up in Paris today' }];
		expectInOrder(await savedParagraph(loaded), [
			'>Meet up in <',
			'<w:smartTag w:element="City">',
			'>Paris<',
			'</w:smartTag>',
			'> today<',
		]);
	});

	it('keeps nested wrappers, a wrapper in a hyperlink, and one at the start of the paragraph', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:smartTag w:element="City">${run('Paris')}</w:smartTag>${run(' and ')}` +
				`<w:hyperlink w:anchor="map"><w:customXml w:element="place"><w:smartTag w:element="City">${run('Lyon')}</w:smartTag>${run(' centre')}</w:customXml></w:hyperlink></w:p>`,
		);
		paragraph.runs = [{ text: 'In Paris or ' }, ...paragraph.runs.filter((item) => item.link)];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'>In <',
			'<w:smartTag w:element="City">',
			'>Paris<',
			'</w:smartTag>',
			'> or <',
		]);
		expectInOrder(element(xml, 'hyperlink'), [
			'<w:customXml w:element="place">',
			'<w:smartTag w:element="City">',
			'>Lyon<',
			'</w:smartTag>',
			'> centre<',
			'</w:customXml>',
		]);
	});

	it('keeps a smart tag and a bookmark when text is typed between them and a replaced character', async () => {
		for (const [open, close] of [
			['<w:smartTag w:element="t">', '</w:smartTag>'],
			['<w:bookmarkStart w:id="1" w:name="Pair"/>', '<w:bookmarkEnd w:id="1"/>'],
		]) {
			const { loaded, paragraph } = await load(
				`<w:p>${run('X')}${open}${run('aa')}${close}${run(' bY')}</w:p>`,
			);
			paragraph.runs = [{ text: 'Zaaa bW' }];
			expectInOrder(await savedParagraph(loaded), ['>Za<', open!, '>aa<', close!, '> bW<']);
		}
	});

	it('keeps an empty smart tag empty when text is typed where it is', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:smartTag w:element="mark"/>${run('B')}</w:p>`,
		);
		paragraph.runs = [{ text: 'AXB' }];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['>AX<', '<w:smartTag w:element="mark"/>', '>B<']);
		expect(xml).not.toContain('annotationEdge');
	});

	it('keeps an empty custom XML element at the start of a hyperlink when other text changes', async () => {
		const empty =
			'<w:customXml w:element="x"><w:customXmlPr><w:attr w:name="id" w:val="42"/></w:customXmlPr></w:customXml>';
		const { loaded, paragraph } = await load(
			`<w:p><w:hyperlink w:anchor="t">${empty}${run('A')}</w:hyperlink>${run('B')}</w:p>`,
		);
		lastOf(paragraph.runs).text = 'C';
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, ['<w:hyperlink w:anchor="t">', empty, '>A<', '</w:hyperlink>', '>C<']);
		expect(xml).not.toContain('annotationEdge');
	});

	it('keeps a smart tag around its text when that text is put in a new hyperlink', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:smartTag w:element="City"><w:smartTagPr><w:attr w:name="n" w:val="1"/></w:smartTagPr>` +
				`${run('Paris')}</w:smartTag>${run(' today')}</w:p>`,
		);
		paragraph.runs = [{ text: 'Paris today', link: { anchor: 't' } }];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'<w:hyperlink w:anchor="t"',
			'<w:smartTag w:element="City">',
			'<w:smartTagPr><w:attr w:name="n" w:val="1"/></w:smartTagPr>',
			'>Paris<',
			'</w:smartTag>',
			'> today<',
			'</w:hyperlink>',
		]);
		expect(xml).not.toContain('annotationEdge');
	});

	it('keeps an empty smart tag at the end of a hyperlink inside it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:hyperlink w:anchor="t">${run('B')}<w:smartTag w:element="x"/></w:hyperlink></w:p>`,
		);
		at(paragraph.runs, 0).text = 'C';
		expectInOrder(await savedParagraph(loaded), [
			'>C<',
			'<w:hyperlink w:anchor="t">',
			'>B<',
			'<w:smartTag w:element="x"/>',
			'</w:hyperlink>',
		]);
	});

	it('keeps a custom XML element around its text when a link is added over part of it', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:customXml w:element="place"><w:customXmlPr><w:attr w:name="id" w:val="1"/></w:customXmlPr>` +
				`${run('Paris')}</w:customXml>${run(' today')}</w:p>`,
		);
		const [paris, today] = [at(paragraph.runs, 0), at(paragraph.runs, 1)];
		paragraph.runs = [
			{ ...paris, text: 'Par', link: { anchor: 't' } },
			{ ...paris, text: 'is' },
			today,
		];
		const saved = await loaded.save();
		const xml =
			(await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')).match(
				/<w:p\b[\s\S]*?<\/w:p>/,
			)?.[0] ?? '';
		expectInOrder(xml, [
			'<w:customXml w:element="place">',
			'<w:customXmlPr><w:attr w:name="id" w:val="1"/></w:customXmlPr>',
			'<w:hyperlink w:anchor="t"',
			'>Par<',
			'</w:hyperlink>',
			'>is<',
			'</w:customXml>',
			'> today<',
		]);
		// The paragraph reads back the same and can be edited again.
		const again = await loadDocx(saved);
		const reread = expectParagraph(again.model.blocks[0]);
		expect(reread.runs.map((item) => [item.text, item.link?.anchor])).toEqual([
			['Par', 't'],
			['is', undefined],
			[' today', undefined],
		]);
		lastOf(reread.runs).text = ' tonight';
		expect(await savedParagraph(again)).toContain('<w:customXml w:element="place">');
	});

	it('keeps nested wrappers around their text when a link is added over part of them', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:customXml w:element="note"><w:smartTag w:element="City">${run('Paris')}</w:smartTag>` +
				`${run(' today')}</w:customXml></w:p>`,
		);
		const [paris, today] = [at(paragraph.runs, 0), at(paragraph.runs, 1)];
		paragraph.runs = [
			{ ...paris, text: 'Par', link: { anchor: 't' } },
			{ ...paris, text: 'is' },
			today,
		];
		const xml = await savedParagraph(loaded);
		expectInOrder(xml, [
			'<w:customXml w:element="note">',
			'<w:smartTag w:element="City">',
			'<w:hyperlink w:anchor="t"',
			'>Par<',
			'</w:hyperlink>',
			'>is<',
			'</w:smartTag>',
			'> today<',
			'</w:customXml>',
		]);
	});

	it('drops a smart tag that would hold a new content control, so the paragraph stays editable', async () => {
		const { loaded, paragraph } = await load(
			`<w:p><w:smartTag w:element="x">${run('A')}${run('B')}${run('C')}</w:smartTag></w:p>`,
		);
		at(paragraph.runs, 1).contentControl = { id: '99', level: 'inline', kind: 'richText' };
		const saved = await loaded.save();
		const again = await loadDocx(saved);
		const reread = expectParagraph(again.model.blocks[0]);
		lastOf(reread.runs).text = 'D';
		const xml = await savedParagraph(again);
		expectInOrder(xml, ['>A<', '<w:sdt>', '>B<', '</w:sdt>', '>D<']);
		expect(xml).not.toContain('<w:smartTag');
	});

	it('keeps a permission range inside the smart tag that held it when other text changes', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:smartTag w:element="x"><w:permStart w:id="1" w:edGrp="everyone"/>${run('B')}` +
				`<w:permEnd w:id="1"/></w:smartTag>${run('C')}</w:p>`,
		);
		at(paragraph.runs, 0).text = 'Z';
		expectInOrder(await savedParagraph(loaded), [
			'>Z<',
			'<w:smartTag w:element="x">',
			'<w:permStart w:id="1" w:edGrp="everyone"/>',
			'>B<',
			'<w:permEnd w:id="1"/>',
			'</w:smartTag>',
			'>C<',
		]);
	});

	it('reads a smart tag inside a simple field as part of the field result', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('Issued ')}<w:fldSimple w:instr=" DOCPROPERTY Issued ">${run('Oct ')}` +
				`<w:smartTag w:element="date">${run('20')}</w:smartTag></w:fldSimple>` +
				`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> by hand</w:t></w:r></w:p>`,
		);
		const result = paragraph.runs.filter((item) => item.field?.simple);
		expect(result.map((item) => item.text)).toEqual(['Oct ', '20']);
		lastOf(paragraph.runs).text = ' by post';
		expectInOrder(await savedParagraph(loaded), [
			'>Issued <',
			'<w:fldSimple w:instr=" DOCPROPERTY Issued ">',
			'>Oct <',
			'>20<',
			'</w:fldSimple>',
			'<w:b/>',
			'> by post<',
		]);
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
			text: '\uF0FE',
			symbol: { font: 'Wingdings', char: 'F0FE' },
		});
		at(paragraph.runs, 0).text = 'Checked ';
		expectInOrder(await savedParagraph(loaded), [
			'>Checked <',
			'<w:sym w:font="Wingdings" w:char="F0FE"/>',
		]);
	});

	it('reads a symbol whose character is half a surrogate pair as unreadable', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:r><w:sym w:font="Wingdings" w:char="D83D"/></w:r></w:p>`,
		);
		expect(at(paragraph.runs, 1)).toMatchObject({ text: '' });
		expect(at(paragraph.runs, 1).symbol).toBeUndefined();
		at(paragraph.runs, 0).text = 'B';
		await expect(loaded.save()).rejects.toThrow('cannot safely relocate');
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

	it('keeps a paragraph read only when a symbol is a character its text can also hold', async () => {
		for (const symbol of [
			// A code outside a symbol font's range is an ordinary character: an en dash or a bullet.
			`<w:r><w:t>a\u2013b</w:t><w:sym w:font="Symbol" w:char="2013"/></w:r>`,
			`<w:r><w:sym w:font="Symbol" w:char="2022"/></w:r>`,
			// The symbol's character in the run's own text could not be told apart from the symbol.
			`<w:r><w:t>\uF0FE</w:t><w:sym w:font="Wingdings" w:char="F0FE"/></w:r>`,
		]) {
			const { loaded, paragraph } = await load(`<w:p>${run('A')}${symbol}</w:p>`);
			expect(at(paragraph.runs, 1).symbol).toBeUndefined();
			at(paragraph.runs, 0).text = 'B';
			await expect(loaded.save()).rejects.toThrow('cannot safely relocate');
		}
	});

	it('still refuses a run that holds two different symbols', async () => {
		const { loaded, paragraph } = await load(
			`<w:p>${run('A')}<w:r><w:sym w:font="Wingdings" w:char="F0FE"/><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>`,
		);
		at(paragraph.runs, 0).text = 'C';
		await expect(loaded.save()).rejects.toThrow('cannot safely relocate');
	});
});
