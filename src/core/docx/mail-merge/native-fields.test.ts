import { expect, it } from 'vitest';
import { plainText } from '../document-stats';
import { previewMailMerge } from './index';
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
		],
	};
	const first = previewMailMerge(model, data, 1);
	expect(first.issues).toEqual([]);
	expect(plainText(first.document)).toBe(
		'Ms. Ada Lovelace\nAcme\n1 Main St\nSeattle, WA 98052\nUSA\nDear Ms. Lovelace,',
	);
	const second = previewMailMerge(model, data, 2);
	expect(second.issues).toEqual([]);
	expect(plainText(second.document)).toBe('Grace\n2 Main St\nSydney\nDear Sir or Madam,');
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
