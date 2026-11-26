import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	HEADER_FOOTER_DESIGNS,
	glossaryBlocksWritable,
	headerFooterDesignBlocks,
	loadDocx,
	parseSignatureLineShape,
	resolveWebVideo,
	type Paragraph,
} from './index';
import { createDocument } from './model';
import { documentPropertyControl } from './quick-parts';
import { saveDocx } from './save';
import { expectParagraph, must } from './test-support/access';
import { parseXml } from './xml';

const context = (properties = {}) => ({
	newId: (() => {
		let n = 0;
		return () => `d${n++}`;
	})(),
	contentWidthTwips: 9360,
	properties: { core: properties, app: {}, custom: [] },
	now: new Date('2026-03-05T10:00:00Z'),
});
const runsOf = (design: (typeof HEADER_FOOTER_DESIGNS)[number], properties = {}) =>
	(headerFooterDesignBlocks(design, context(properties))[0] as Paragraph).runs;

describe('header and footer designs', () => {
	it('uses only PAGE fields and shows controls with placeholders, for every design', () => {
		for (const design of HEADER_FOOTER_DESIGNS) {
			const runs = runsOf(design);
			const instructions = runs.flatMap((run) => (run.field ? [run.field.instr] : []));
			expect(instructions, design).toEqual(
				design === 'authorAndPage' || design === 'dateAndPage' ? [' PAGE '] : [],
			);
			for (const run of runs) expect(run.text.includes('@')).toBe(false);
		}
		expect(runsOf('title')[0]).toMatchObject({
			text: '[Document title]',
			contentControl: { alias: 'Title', showingPlaceholder: true },
		});
		expect(runsOf('titleAndAuthor', { title: 'Plan', creator: 'Ada' }).map((r) => r.text)).toEqual([
			'Plan',
			'\t',
			'\t',
			'Ada',
		]);
		expect(runsOf('authorAndPage')[0]).toMatchObject({ text: '[Author name]' });
	});

	it('fills the date picker with today and writes a w:date control', async () => {
		const date = runsOf('dateAndPage')[0]!;
		expect(date.text).toBe('March 5, 2026');
		expect(date.contentControl).toMatchObject({ kind: 'date', dateFormat: 'MMMM d, yyyy' });
		const model = createDocument();
		model.sections = [
			{
				type: 'nextPage',
				orientation: 'portrait',
				pageWidthTwips: 12240,
				pageHeightTwips: 15840,
				marginTopTwips: 1440,
				marginLeftTwips: 1440,
				marginRightTwips: 1440,
				marginBottomTwips: 1440,
				columns: { count: 1, equalWidth: true },
				headers: {
					default: {
						partName: 'word/header1.xml',
						blocks: headerFooterDesignBlocks('dateAndPage', context()),
					},
				},
			},
		] as never;
		const zip = await JSZip.loadAsync(await saveDocx(model));
		const header = await must(zip.file('word/header1.xml')).async('string');
		expect(header).toContain('<w:date w:fullDate="2026-03-05T00:00:00Z">');
		expect(header).toContain('PAGE');
	});
});

describe('glossary writing', () => {
	const entry = (blocks: Paragraph[], name = 'Part', gallery = 'docParts') => ({
		name,
		gallery,
		category: 'General',
		blocks,
	});
	const plain = (text: string): Paragraph => ({ type: 'paragraph', id: 'p', runs: [{ text }] });

	it('refuses a hyperlink, a note and a comment range, which need relationships', () => {
		const link = { text: 'x', link: { href: 'https://example.com' } } as never;
		expect(glossaryBlocksWritable([{ ...plain('x'), runs: [link] }])).toBe(false);
		expect(
			glossaryBlocksWritable([
				{ ...plain('x'), runs: [{ text: '', noteReference: { kind: 'footnote', id: '1' } }] },
			]),
		).toBe(false);
		expect(glossaryBlocksWritable([plain('plain')])).toBe(true);
	});

	it('writes valid docPart types and replaces an entry of the same key', async () => {
		const model = createDocument();
		model.glossary = [entry([plain('one')], 'Note', 'autoTxt'), entry([plain('A')], 'Same')];
		const first = await loadDocx(await saveDocx(model));
		const xml = await must(
			(await JSZip.loadAsync(await first.save(first.model))).file('word/glossary/document.xml'),
		).async('string');
		expect(xml).toContain('<w:type w:val="normal"/>');
		expect(xml).not.toContain('w:val="autoTxt"/></w:types>');
		const next = first.model;
		next.glossary = [
			...must(next.glossary).filter((e) => e.name !== 'Same'),
			entry([plain('B')], 'Same'),
		];
		const out = await must(
			(await JSZip.loadAsync(await first.save(next))).file('word/glossary/document.xml'),
		).async('string');
		expect(out.match(/w:val="Same"/g)).toHaveLength(1);
		expect(out).toContain('>B<');
		expect(out).not.toContain('>A<');
	});

	it('writes into the glossary part the document already names, with one relationship', async () => {
		const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
		zip.file(
			'word/glossaryX/doc.xml',
			'<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>',
		);
		zip.file(
			'word/_rels/document.xml.rels',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossaryX/doc.xml"/></Relationships>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		loaded.model.glossary = [entry([plain('Kept')])];
		const saved = await JSZip.loadAsync(await loaded.save(loaded.model));
		expect(await must(saved.file('word/glossaryX/doc.xml')).async('string')).toContain('Kept');
		expect(saved.file('word/glossary/document.xml')).toBeNull();
		const rels = await must(saved.file('word/_rels/document.xml.rels')).async('string');
		expect(rels.match(/glossaryDocument/g)).toHaveLength(1);
	});
});

describe('bound document properties', () => {
	it('creates a missing app part for a Company control', async () => {
		const loaded = await loadDocx(await saveDocx(createDocument()));
		const { control } = must(documentPropertyControl('company', new Set(), 'Société'));
		expect(control.alias).toBe('Société');
		expect(must(control.sourceXml)).toContain('w:val="Société"');
		expectParagraph(loaded.model.blocks[0]).runs = [{ text: 'Contoso', contentControl: control }];
		const zip = await JSZip.loadAsync(await loaded.save(loaded.model));
		expect(await must(zip.file('docProps/app.xml')).async('string')).toContain('Contoso');
		expect(await must(zip.file('[Content_Types].xml')).async('string')).toContain(
			'/docProps/app.xml',
		);
		expect(await must(zip.file('_rels/.rels')).async('string')).toContain('extended-properties');
	});

	it('keeps a Company control edit together with an edited Manager of the package information', async () => {
		const loaded = await loadDocx(await saveDocx(createDocument()));
		const model = loaded.model;
		const { control } = must(documentPropertyControl('company', new Set()));
		expectParagraph(model.blocks[0]).runs = [{ text: 'Contoso', contentControl: control }];
		model.packageProperties = { core: {}, app: { manager: 'Grace' }, custom: [] };
		const zip = await JSZip.loadAsync(await loaded.save(model));
		const app = await must(zip.file('docProps/app.xml')).async('string');
		expect(app).toContain('Contoso');
		expect(app).toContain('Company');
	});
});

describe('signature line attributes', () => {
	const parse = (attributes: string) => {
		const doc = parseXml(
			`<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><o:signatureline id="{A}" ${attributes}/></v:shape>`,
		);
		return parseSignatureLineShape(doc.documentElement);
	};
	it('defaults a missing showsigndate to true and allowcomments to false', () => {
		expect(parse('')).toMatchObject({ showSignDate: true, allowComments: false });
		expect(parse('showsigndate="f" allowcomments="t"')).toMatchObject({
			showSignDate: false,
			allowComments: true,
		});
	});
	it('reads instructions flagged by an unqualified signinginstructionsset', () => {
		expect(parse('signinginstructionsset="t" o:signinginstructions="Sign"')?.instructions).toBe(
			'Sign',
		);
	});
});

describe('online video embeds', () => {
	const allow = ['video.example.com'];
	it('refuses script and data addresses, lookalike hosts, user info and data-src', () => {
		for (const input of [
			'javascript:alert(1)',
			'data:text/html,<b>x</b>',
			'https://video.example.com.evil.org/embed/1',
			'https://evilvideo.example.com.evil.org/x',
			'https://user:pw@video.example.com/embed/1',
			'https://video.example.com@evil.org/embed/1',
			'<iframe data-src="https://video.example.com/embed/1"></iframe>',
		])
			expect(resolveWebVideo(input, allow).ok, input).toBe(false);
		expect(resolveWebVideo('https://evilvideo.example.com/x', allow)).toMatchObject({ ok: false });
		expect(resolveWebVideo('https://a.video.example.com/embed/1', allow).ok).toBe(true);
	});
});
