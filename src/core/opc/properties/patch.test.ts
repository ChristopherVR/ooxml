import { describe, expect, it } from 'vitest';

import { NS } from '../../xml/index';
import {
	STRICT_APP_NAMESPACES,
	customPropertyFromText,
	parseAppProperties,
	parseCoreProperties,
	replaceTitleGroup,
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
} from './index';

const CORE_HEAD = `<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}"`;

describe('byte-preserving property patches', () => {
	it('keeps comments, indentation and empty-element forms the edit does not touch', () => {
		const source = `${CORE_HEAD}>\n  <!-- a > b -->\n  <dc:title/>\n  <dc:subject a='x>y'></dc:subject>\n</cp:coreProperties>`;
		const out = writeCoreProperties({ creator: 'Ann' }, source);
		expect(out).toBe(
			`${CORE_HEAD}>\n  <!-- a > b -->\n  <dc:title/>\n  <dc:subject a='x>y'></dc:subject>\n<dc:creator>Ann</dc:creator></cp:coreProperties>`,
		);
	});

	it('fills a self-closing element and keeps its attributes', () => {
		const source = `${CORE_HEAD}><dc:title xml:lang="en"/></cp:coreProperties>`;
		expect(writeCoreProperties({ title: 'A\rB' }, source)).toContain(
			'<dc:title xml:lang="en">A&#xD;B</dc:title>',
		);
	});

	it('types a written date and declares the xsi and dcterms prefixes it needs', () => {
		const source = `${CORE_HEAD}><dc:title>T</dc:title></cp:coreProperties>`;
		const out = writeCoreProperties({ title: 'T', modified: '2026-10-08T09:30:00Z' }, source);
		expect(out).toContain(
			`${CORE_HEAD} xmlns:dcterms="${NS.dcterms}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`,
		);
		expect(out).toContain(
			'<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-08T09:30:00Z</dcterms:modified>',
		);
		expect(parseCoreProperties(out).modified).toBe('2026-10-08T09:30:00Z');
	});

	it('adds the missing type to an existing untyped date it rewrites', () => {
		const source = `${CORE_HEAD} xmlns:dcterms="${NS.dcterms}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created>2020-01-01T00:00:00Z</dcterms:created></cp:coreProperties>`;
		expect(writeCoreProperties({ created: '2021-01-01T00:00:00Z' }, source)).toContain(
			'<dcterms:created xsi:type="dcterms:W3CDTF">2021-01-01T00:00:00Z</dcterms:created>',
		);
	});

	it('opens a self-closing root to add children', () => {
		const source = `<?xml version="1.0" encoding="UTF-8"?>\n<Properties xmlns="${NS.extendedProperties}" xmlns:vt="${NS.vt}"/>`;
		expect(writeAppProperties({ slides: 2, words: 5 }, source)).toBe(
			`<?xml version="1.0" encoding="UTF-8"?>\n<Properties xmlns="${NS.extendedProperties}" xmlns:vt="${NS.vt}"><Words>5</Words><Slides>2</Slides></Properties>`,
		);
	});

	it('declares vt when a vector is written into a part that lacks it', () => {
		const source = `<Properties xmlns="${NS.extendedProperties}"><Slides>1</Slides></Properties>`;
		const out = writeAppProperties({ slides: 1, titlesOfParts: ['A'] }, source);
		expect(out).toContain(`<Properties xmlns="${NS.extendedProperties}" xmlns:vt="${NS.vt}">`);
		expect(out).toContain(
			'<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>A</vt:lpstr></vt:vector></TitlesOfParts>',
		);
	});

	it('reads and patches a Strict part in its own namespaces', () => {
		const ns = STRICT_APP_NAMESPACES;
		const source = `<Properties xmlns="${ns.extendedProperties}" xmlns:v="${ns.vt}"><Slides>1</Slides><TitlesOfParts><v:vector size="1" baseType="lpstr"><v:lpstr>One</v:lpstr></v:vector></TitlesOfParts></Properties>`;
		expect(parseAppProperties(source)).toEqual({ slides: 1, titlesOfParts: ['One'] });
		const out = writeAppProperties({ slides: 2, titlesOfParts: ['One', 'Two'] }, source);
		expect(out).toContain('<Slides>2</Slides>');
		expect(out).toContain(
			'<v:vector size="2" baseType="lpstr"><v:lpstr>One</v:lpstr><v:lpstr>Two</v:lpstr></v:vector>',
		);
		expect(parseAppProperties(out)).toEqual({ slides: 2, titlesOfParts: ['One', 'Two'] });
	});
});

describe('title groups', () => {
	const pairs = [
		{ name: 'Fonts Used', count: 2 },
		{ name: 'Theme', count: 1 },
		{ name: 'Slide Titles', count: 2 },
	];
	const titles = ['Arial', 'Calibri', 'Office Theme', 'Old 1', 'Old 2'];

	it('replaces one group and keeps the others', () => {
		expect(
			replaceTitleGroup(pairs, titles, { name: 'Slide Titles', titles: ['A', 'B', 'C'] }),
		).toEqual({
			headingPairs: [
				{ name: 'Fonts Used', count: 2 },
				{ name: 'Theme', count: 1 },
				{ name: 'Slide Titles', count: 3 },
			],
			titlesOfParts: ['Arial', 'Calibri', 'Office Theme', 'A', 'B', 'C'],
		});
	});

	it('appends the group when no group matches', () => {
		const out = replaceTitleGroup(pairs.slice(0, 2), titles.slice(0, 3), {
			name: 'Slide Titles',
			titles: ['A'],
			match: (name) => name.toLowerCase().includes('slide'),
		});
		expect(out.headingPairs.at(-1)).toEqual({ name: 'Slide Titles', count: 1 });
		expect(out.titlesOfParts).toEqual(['Arial', 'Calibri', 'Office Theme', 'A']);
	});
});

describe('custom properties from text', () => {
	it('types a value only when its text survives', () => {
		expect(customPropertyFromText('a', 'i4', '42')).toEqual({ name: 'a', type: 'i4', value: 42 });
		expect(customPropertyFromText('a', 'i4', '042')).toEqual({
			name: 'a',
			type: 'raw',
			xml: '<vt:i4>042</vt:i4>',
		});
		expect(customPropertyFromText('a', 'r8', '2.5')).toEqual({ name: 'a', type: 'r8', value: 2.5 });
		expect(customPropertyFromText('a', 'bool', '1')).toEqual({
			name: 'a',
			type: 'raw',
			xml: '<vt:bool>1</vt:bool>',
		});
		expect(customPropertyFromText('a', 'ui4', '7')).toEqual({
			name: 'a',
			type: 'raw',
			xml: '<vt:ui4>7</vt:ui4>',
		});
		expect(customPropertyFromText('a', 'no such', 'x')).toEqual({
			name: 'a',
			type: 'lpwstr',
			value: 'x',
		});
	});

	it('writes the text exactly', () => {
		const xml = writeCustomProperties([customPropertyFromText('Id', 'lpstr', 'R&D')]);
		expect(xml).toContain('<vt:lpstr>R&amp;D</vt:lpstr>');
	});
});
