import JSZip from 'jszip';
import { expect, test } from 'vitest';
import {
	createDocument,
	loadDocx,
	saveDocx,
	type DocumentModel,
	type Paragraph,
	type TextRun,
} from '../index';
import { sectionsOf } from '../section-layout';
import { twips } from '../units';
import {
	authorityCode,
	authorityEntryCode,
	buildReferenceTable,
	citationCode,
	citationEvaluator,
	editCitationCode,
	findReferenceTables,
	indexCode,
	indexEntryCode,
	indexSection,
	readBibliographyXml,
	referenceFieldRuns,
	refreshCitations,
	writeBibliographyXml,
	type Bibliography,
} from './index';

const paragraph = (id: string, ...runs: TextRun[]): Paragraph => ({ type: 'paragraph', id, runs });
const mark = (main: string, extra: Partial<Parameters<typeof indexEntryCode>[0]> = {}) =>
	referenceFieldRuns(
		indexEntryCode({
			main,
			subentry: '',
			crossReference: '',
			bookmark: '',
			bold: false,
			italic: false,
			...extra,
		}),
	);
const text = (paragraphs: Paragraph[]) =>
	paragraphs.map(({ runs }) => runs.map((run) => run.text).join(''));
const index = (runIn = false) =>
	indexCode({ columns: 2, runIn, rightAlign: true, leader: 'hyphen', lcid: 1053 });

test('builds indented and run-in indexes with localized grouping, range, and page formatting', () => {
	const model: DocumentModel = {
		...createDocument(),
		blocks: [
			paragraph('first', ...mark('Zebra', { subentry: 'A:B', bold: true, italic: true })),
			paragraph('second', ...mark('Ål', { bookmark: 'range' })),
			paragraph('third', ...mark('Ångström', { crossReference: 'See Zebra' })),
			{ ...paragraph('start', { text: 'Open' }), bookmarks: ['range'] },
			paragraph('end', { text: 'Close' }),
		],
		sourceBookmarkEnds: [['range', 'end']],
	};
	const pages = new Map([
		['first', '2'],
		['start', '3'],
		['end', '5'],
	]);
	const options = { empty: 'None', leader: 'hyphen' as const, page: (id: string) => pages.get(id) };
	const indented = buildReferenceTable(model, index(), options);
	expect(text(indented)).toEqual(['Z', 'Zebra', 'A:B\t2', 'Å', 'Ål\t3–5', 'Ångström. See Zebra']);
	expect(indented[2]?.runs.find((run) => run.text === '2')).toMatchObject({
		bold: true,
		italic: true,
	});
	expect(indented[2]).toMatchObject({ tabStops: [{ align: 'right', leader: 'hyphen' }] });
	expect(text(buildReferenceTable(model, index(true), options))).toContain('Zebra: A:B 2');
});

test('uses the index section width and inserts one column inside a two-column section', async () => {
	const initial = createDocument();
	const section = sectionsOf(initial)[0]!;
	const model: DocumentModel = {
		...initial,
		blocks: [
			paragraph('first', { text: 'Before the two-column section' }),
			paragraph('entry', ...mark('Zebra')),
			paragraph('insert'),
			paragraph('after', { text: 'Stay in two columns' }),
		],
		sections: [
			{ ...section, endsAtBlockId: 'first' },
			{
				...section,
				endsAtBlockId: 'after',
				pageWidthTwips: twips(14_400),
				marginLeftTwips: twips(1_000),
				marginRightTwips: twips(2_000),
				columns: { count: 2, equalWidth: true },
			},
		],
	};
	const instruction = indexCode({
		columns: 1,
		runIn: false,
		rightAlign: true,
		leader: 'dot',
		lcid: 1033,
	});
	const options = { empty: 'None', page: () => '2', blockIndex: 2 };
	const output = buildReferenceTable(model, instruction, options).map((p, i) => ({
		...p,
		id: `index-${i}`,
	}));
	const edit = indexSection(model, 2, 2, output, 1, () => 'new-section');
	expect(edit?.sections.map((item) => item.columns.count)).toEqual([1, 2, 1, 2]);
	const saved = await saveDocx({
		...model,
		blocks: [model.blocks[0]!, model.blocks[1]!, ...output, model.blocks[3]!],
		sections: edit!.sections,
	});
	const reopened = (await loadDocx(saved)).model;
	expect(reopened.sections?.map((item) => item.columns.count)).toEqual([1, 2, 1, 2]);
	const [location] = findReferenceTables(reopened.blocks, 'INDEX');
	expect(location?.start).toBe(2);
	const updated = buildReferenceTable(reopened, location!.instruction, options);
	expect(updated[1]?.tabStops).toEqual([{ posTwips: 11_400, align: 'right', leader: 'dot' }]);
	expect(text(updated)).toEqual(['Z', 'Zebra\t2']);
});

test('filters authorities by category, joins short citations, and applies passim and original formatting', () => {
	const long = 'Doe v. City, 42 F. 1';
	const uri = authorityEntryCode(long, 'Doe', 1);
	const model: DocumentModel = {
		...createDocument(),
		blocks: Array.from({ length: 5 }, (_, index) =>
			paragraph(
				String(index),
				...referenceFieldRuns(index === 0 ? uri : 'TA \\s "Doe" \\c 1').map((run) =>
					run.fieldCode ? { ...run, italic: true } : run,
				),
			),
		),
	};
	const options = { empty: 'None', page: (id: string) => String(Number(id) + 1) };
	const kept = buildReferenceTable(model, authorityCode(1, true, true), options);
	expect(text(kept)).toEqual([`${long}\tpassim`]);
	expect(kept[0]?.runs.find((run) => run.text === long)).toMatchObject({ italic: true });
	const plain = buildReferenceTable(model, authorityCode(1, false, false), options);
	expect(text(plain)).toEqual([`${long}\t1, 2, 3, 4, 5`]);
	expect(plain[0]?.runs.find((run) => run.text === long)?.italic).toBeUndefined();
	expect(text(buildReferenceTable(model, authorityCode(2, true, true), options))).toEqual(['None']);
});

const bibliography: Bibliography = {
	style: 'APA',
	sources: [
		{
			tag: 'Doe2020',
			type: 'Book',
			fields: { Title: 'Field studies', Year: '2020', Publisher: 'Example Press' },
			author: { kind: 'people', people: [{ first: 'Jane', middle: '', last: 'Doe' }] },
		},
		{ tag: 'Later', type: 'Book', fields: {}, placeholder: true },
	],
};
const selectors = { pages: '12', suppressAuthor: false, suppressTitle: false, suppressYear: false };
const cite = citationCode('Doe2020', 1033, selectors);

test('updates citation results with style and suppression, and keeps locked fields', () => {
	const evaluate = citationEvaluator(bibliography);
	expect(evaluate(cite)).toBe('(Doe, 2020, p. 12)');
	expect(citationEvaluator({ ...bibliography, style: 'MLA' })(cite)).toBe('(Doe 12)');
	expect(citationEvaluator({ ...bibliography, style: 'Chicago' })(cite)).toBe('(Doe 2020, 12)');
	const edited = editCitationCode(cite + ' \\f "See "', {
		...selectors,
		suppressAuthor: true,
		suppressYear: true,
	});
	expect(edited && evaluate(edited)).toBe('(See Field studies, p. 12)');
	expect(evaluate(citationCode('Later', 1033, selectors))).toBe('([Later])');
	const original = paragraph('p', ...referenceFieldRuns(cite, [{ text: 'cached', italic: true }]));
	expect(refreshCitations(original, evaluate).runs[3]).toMatchObject({
		italic: true,
		text: '(Doe, 2020, p. 12)',
	});
	const locked = {
		...original,
		runs: original.runs.map((run) =>
			run.fieldChar === 'begin' ? { ...run, fieldFlags: { locked: true } } : run,
		),
	};
	expect(refreshCitations(locked, evaluate)).toBe(locked);
});

test('saves all reference fields, continuous index columns and the Word source package', async () => {
	const initial = createDocument();
	const seed = {
		...initial,
		blocks: [
			paragraph(
				'p-1',
				{ text: 'Zebra. ' },
				...mark('Zebra'),
				...referenceFieldRuns(authorityEntryCode('Doe v. City', 'Doe', 1)),
				...referenceFieldRuns(cite, [{ text: '(Doe, 2020, p. 12)' }]),
			),
			paragraph('empty'),
		],
	};
	const options = { empty: 'None', bibliography, page: () => '1' };
	const identified = (paragraphs: Paragraph[], prefix: string) =>
		paragraphs.map((item, i) => ({ ...item, id: prefix + i }));
	const columns = indexSection(
		seed,
		1,
		1,
		identified(buildReferenceTable(seed, index(), options), 'index-'),
		2,
		() => 'after-index',
	);
	expect(columns?.sections.map((item) => item.columns.count)).toEqual([1, 2, 1]);
	const model: DocumentModel = {
		...seed,
		blocks: [
			seed.blocks[0]!,
			...columns!.paragraphs,
			...identified(buildReferenceTable(seed, 'BIBLIOGRAPHY', options), 'bib-'),
			...identified(buildReferenceTable(seed, authorityCode(1, false, true), options), 'toa-'),
		],
		sections: columns!.sections,
		packageInformation: { bibliography: { xml: writeBibliographyXml(bibliography) } },
	};
	const bytes = await saveDocx(model);
	const reopened = await loadDocx(bytes);
	const part = reopened.model.packageInformation?.bibliography;
	expect(part?.part).toBe('customXml/item1.xml');
	expect(part && readBibliographyXml(part.xml)).toEqual(bibliography);
	expect(reopened.model.sections?.map((item) => item.columns.count)).toEqual([1, 2, 1]);
	for (const name of ['INDEX', 'BIBLIOGRAPHY', 'TOA'] as const)
		expect(findReferenceTables(reopened.model.blocks, name)).toHaveLength(1);
	const zip = await JSZip.loadAsync(bytes);
	expect(await zip.file('customXml/itemProps1.xml')?.async('string')).toContain('bibliography');
	expect(await zip.file('word/_rels/document.xml.rels')?.async('string')).toContain(
		'../customXml/item1.xml',
	);
	expect(await zip.file('word/document.xml')?.async('string')).toMatch(/CITATION .*Doe2020/);
	expect(await reopened.save()).toEqual(bytes);
});

test('imports a bookmark spanning paragraphs and retains opaque source elements after a source edit', async () => {
	const zip = new JSZip();
	const xml = writeBibliographyXml(bibliography).replace(
		'</b:Source>',
		'<b:Comments>Keep</b:Comments></b:Source>',
	);
	zip.file(
		'word/document.xml',
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="0" w:name="chapter"/><w:r><w:t>Start</w:t></w:r></w:p><w:p><w:r><w:t>End</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p></w:body></w:document>',
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/sources.xml"/></Relationships>',
	);
	zip.file('customXml/sources.xml', xml);
	const reopened = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	expect(reopened.model.sourceBookmarkEnds).toEqual([['chapter', 'p1']]);
	const info = reopened.model.packageInformation;
	expect(info?.bibliography?.xml).toBe(xml);
	const changed = writeBibliographyXml({ ...bibliography, style: 'MLA' }, xml);
	const next = await reopened.save({
		...reopened.model,
		packageInformation: { ...info, bibliography: { part: 'customXml/sources.xml', xml: changed } },
	});
	const savedPart = (await loadDocx(next)).model.packageInformation?.bibliography;
	expect(savedPart?.xml).toContain('<b:Comments>Keep</b:Comments>');
	expect(savedPart?.xml).toContain('StyleName="MLA"');
});
