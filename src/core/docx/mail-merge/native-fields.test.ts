import { expect, it } from 'vitest';
import { plainText } from '../document-stats';
import { createDocument } from '../model';
import { loadDocx } from '../parse';
import { saveDocx } from '../save';
import { mailMergeAddressBlock, mailMergeGreetingLine, previewMailMerge } from './index';
import { code, complex, paragraph, simple, template } from './test-fixture';

const escapeTemplate = (instruction: string) =>
	instruction.replaceAll('<', '&lt;').replaceAll('>', '&gt;');
const field = (instruction: string, cached?: string) =>
	complex(code(escapeTemplate(instruction)), cached);

// Captured from Word 16's Insert Address Block and Insert Greeting Line dialogs via Field.Code.Text.
const wordAddress =
	' ADDRESSBLOCK \\f "<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>><< _SUFFIX0_>>\r<<_COMPANY_\r>><<_STREET1_\r>><<_STREET2_\r>><<_CITY_>><<, _STATE_>><< _POSTAL_>><<\r_COUNTRY_>>" \\l 1033 \\c 2 \\e "Australia" \\d ';
const wordGreeting =
	' GREETINGLINE \\f "<<_BEFORE_ Dear >><<_TITLE0_>><< _LAST0_>>\n<<_AFTER_ ,>>" \\l 1033 \\e "Dear Sir or Madam," ';

it('renders the default Word 16 codes with multiline groups and literal spaces', async () => {
	const addressRuns = escapeTemplate(wordAddress)
		.split('\r')
		.map((part, index) => (index ? code(part).replace('<w:r>', '<w:r><w:cr/>') : code(part)))
		.join('');
	const { model } = await template(
		paragraph(complex(addressRuns)) +
			paragraph(simple(escapeTemplate(wordGreeting).replace('\n', '&#xA;'), 'GreetingLine')),
	);
	const data = {
		records: [
			{
				CourtesyTitle: 'Ms.',
				FirstName: 'Ada',
				LastName: 'Lovelace',
				Company: 'Acme',
				Address1: '1 Main St',
				City: 'Seattle',
				State: 'WA',
				PostalCode: '98052',
				CountryorRegion: 'USA',
			},
			{ FirstName: 'Grace', Address1: '2 Main St', City: 'Sydney', CountryorRegion: 'Australia' },
			{},
			{ Company: 'Company Only', CountryorRegion: 'Australia' },
			{ LastName: 'Only Last', CountryorRegion: 'Australia' },
			{ CourtesyTitle: 'Dr.' },
		],
	};
	const first = previewMailMerge(model, data, 1);
	expect(first.issues).toEqual([]);
	expect(plainText(first.document)).toBe(
		'Ms. Ada Lovelace\nAcme\n1 Main St\nSeattle, WA 98052\nUSA\nDear Ms. Lovelace,',
	);
	const second = previewMailMerge(model, data, 2);
	expect(second.issues).toEqual([]);
	expect(plainText(second.document)).toBe('Grace\n2 Main St\nSydney\nDear Grace,');
	expect(plainText(previewMailMerge(model, data, 3).document).trim()).toBe('Dear Sir or Madam,');
	expect(plainText(previewMailMerge(model, data, 4).document)).toBe(
		'Company Only\n\nDear Sir or Madam,',
	);
	expect(plainText(previewMailMerge(model, data, 5).document)).toBe(
		' Only Last\nDear Sir or Madam,',
	);
	expect(plainText(previewMailMerge(model, data, 6).document)).toBe('Dr. \nDear Sir or Madam,');
});

it('writes the captured Word templates with decimal language IDs through a DOCX round trip', async () => {
	const address = mailMergeAddressBlock({
		country: '2',
		excludeCountry: 'Australia',
		formatByCountry: true,
	});
	const greeting = mailMergeGreetingLine();
	const document = createDocument();
	document.blocks = [
		{ type: 'paragraph', id: 'address', runs: address },
		{ type: 'paragraph', id: 'greeting', runs: greeting },
	];
	const { model } = await loadDocx(await saveDocx(document));
	const [addressBlock, greetingBlock] = model.blocks;
	expect(
		addressBlock?.type === 'paragraph' &&
			addressBlock.runs.map((run) => run.fieldCode ?? '').join(''),
	).toBe(wordAddress);
	expect(greetingBlock?.type === 'paragraph' && greetingBlock.runs[0]?.field).toEqual({
		instr: wordGreeting.trim(),
		simple: true,
	});
	expect(mailMergeAddressBlock({ language: 'fr-FR' }).at(-2)?.field?.instr).toContain('\\l 1036');
	expect(mailMergeGreetingLine({ language: '0407' })[0]?.field?.instr).toContain('\\l 1031');
});

it('emits a group only when all of its placeholders have values and adds no spaces', async () => {
	const { model } = await template(
		paragraph(field(' ADDRESSBLOCK \\f "<<_FIRST0_>><<( _MIDDLE0_\n _SUFFIX0_)>><<_LAST0_>>" ')),
	);
	const data = {
		records: [
			{ FirstName: 'Ada', MiddleName: 'Byron', LastName: 'Lovelace' },
			{ FirstName: 'Ada', MiddleName: 'Byron', Suffix: 'II', LastName: 'Lovelace' },
		],
	};
	expect(plainText(previewMailMerge(model, data, 1).document)).toBe('AdaLovelace');
	expect(plainText(previewMailMerge(model, data, 2).document)).toBe('Ada( Byron\n II)Lovelace');
});
