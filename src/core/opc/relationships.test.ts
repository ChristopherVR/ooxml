import { describe, expect, it } from 'vitest';

import { NS } from '../xml/index';
import { canonicalRelationshipsXml, parseRelationshipsPart } from './relationships';

const TYPE = 'urn:test/type';
const part = (body: string, rootAttributes = '') =>
	`<Relationships xmlns="${NS.rels}"${rootAttributes}>${body}</Relationships>`;
const relationship = (attributes: string) => `<Relationship ${attributes}/>`;
const VALID = relationship(`Id="rId1" Type="${TYPE}" Target="a.xml"`);

describe('parseRelationshipsPart', () => {
	it('reads relationships in document order with the default target mode', () => {
		const xml = part(
			VALID + relationship(`Id="rId0" Type="${TYPE}" Target="http://x/" TargetMode="External"`),
		);
		expect(parseRelationshipsPart(xml)).toEqual([
			{ id: 'rId1', type: TYPE, target: 'a.xml', mode: 'Internal' },
			{ id: 'rId0', type: TYPE, target: 'http://x/', mode: 'External' },
		]);
	});

	it('accepts the relationships namespace under a prefix', () => {
		const xml =
			`<r:Relationships xmlns:r="${NS.rels}">` +
			`<r:Relationship Id="rId1" Type="${TYPE}" Target="a.xml"/></r:Relationships>`;
		expect(parseRelationshipsPart(xml)).toHaveLength(1);
	});

	it('reads a part that starts with a byte order mark', () => {
		expect(parseRelationshipsPart(`\uFEFF${part(VALID)}`)).toHaveLength(1);
	});

	it('accepts an empty Target, which xsd:anyURI allows', () => {
		const xml = part(relationship(`Id="rId1" Type="${TYPE}" Target="" TargetMode="External"`));
		expect(parseRelationshipsPart(xml)).toEqual([
			{ id: 'rId1', type: TYPE, target: '', mode: 'External' },
		]);
	});

	it.each(['rId1 ', ' rId1', '  rId1  '])('refuses Id %j beside rId1', (id) => {
		const xml = part(VALID + relationship(`Id="${id}" Type="${TYPE}" Target="b.xml"`));
		expect(() => parseRelationshipsPart(xml)).toThrow(/duplicate Id/u);
	});

	it('keeps Ids that differ only in case apart, as xsd:ID does', () => {
		const xml = part(VALID + relationship(`Id="RID1" Type="${TYPE}" Target="b.xml"`));
		expect(parseRelationshipsPart(xml).map(({ id }) => id)).toEqual(['rId1', 'RID1']);
	});

	it.each([
		['a root in another namespace', `<Relationships xmlns="urn:x">${VALID}</Relationships>`],
		['an attribute on the root', part(VALID, ' Extra="1"')],
		[
			'a child in another namespace',
			part(`${VALID}<Relationship xmlns="urn:x" Id="rId2" Type="${TYPE}" Target="b.xml"/>`),
		],
		[
			'a prefixed child in another namespace',
			part(`${VALID}<x:Relationship xmlns:x="urn:x" Id="rId2" Type="${TYPE}" Target="b.xml"/>`),
		],
		['a child with no namespace', part(`<Relationship xmlns="" Id="r" Type="t" Target="b"/>`)],
		['another child element', part(`${VALID}<Other/>`)],
		['a repeated Id', part(VALID + VALID)],
		['a missing Target', part(relationship(`Id="rId1" Type="${TYPE}"`))],
		['a missing Type', part(relationship('Id="rId1" Target="a.xml"'))],
		['a missing Id', part(relationship(`Type="${TYPE}" Target="a.xml"`))],
		[
			'a target mode in another case',
			part(relationship(`Id="r" Type="${TYPE}" Target="a" TargetMode="external"`)),
		],
		['an unknown attribute', part(relationship(`Id="r" Type="t" Target="a" Other="1"`))],
		['a repeated attribute', part(relationship(`Id="r" Type="t" Target="a" Target="b"`))],
		['a document type declaration', `<!DOCTYPE x []>${part(VALID)}`],
	])('rejects %s', (_label, xml) => {
		expect(() => parseRelationshipsPart(xml)).toThrow();
	});
});

describe('canonicalRelationshipsXml', () => {
	it('rewrites a part to the relationships the strict reader accepted', () => {
		const xml = part(`<!-- note -->\n  ${relationship(`Id="r&#49;" Type="t&amp;u" Target="a"`)}`);
		expect(parseRelationshipsPart(canonicalRelationshipsXml(xml))).toEqual([
			{ id: 'r1', type: 't&u', target: 'a', mode: 'Internal' },
		]);
	});

	it('keeps an empty Target and both spellings of a case-distinct Id', () => {
		const xml = part(
			relationship(`Id="a" Type="t" Target=""`) + relationship(`Id="A" Type="t" Target="b"`),
		);
		expect(parseRelationshipsPart(canonicalRelationshipsXml(xml))).toEqual([
			{ id: 'a', type: 't', target: '', mode: 'Internal' },
			{ id: 'A', type: 't', target: 'b', mode: 'Internal' },
		]);
	});
});
