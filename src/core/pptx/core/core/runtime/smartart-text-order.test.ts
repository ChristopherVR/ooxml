import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../../types';
import {
	annotateSmartArtTextOrder,
	orderedSmartArtTextEntries,
	smartArtChildOrder,
} from './smartart-text-order';
import { growth } from '../../../../test-support/growth';

describe('smartArt text order annotation', () => {
	it.each(['p:shape', 'p-extension', 'p.extra'])(
		'does not mistake <%s> for an unprefixed paragraph',
		(wrapper) => {
			const paragraph: XmlObject = {
				r: [{ t: 'First' }, { t: 'Second' }],
				tab: '',
			};
			annotateSmartArtTextOrder(
				`<${wrapper}><txBody><p><r><t>First</t></r><tab/><r><t>Second</t></r></p></txBody></${wrapper}>`,
				{ txBody: { p: paragraph } },
			);
			expect(smartArtChildOrder(paragraph)).toStrictEqual(['r', 'tab', 'r']);
			expect(orderedSmartArtTextEntries(paragraph).map(([key]) => key)).toStrictEqual([
				'r',
				'tab',
				'r',
			]);
		},
	);

	it.each(['r', 'fld', 'br'])(
		'does not mistake the %s namespace prefix for a text item',
		(name) => {
			const item: XmlObject = { rPr: {} };
			const paragraph: XmlObject = { [name]: item };
			annotateSmartArtTextOrder(
				`<${name}:wrapper><txBody><p><${name}><rPr/></${name}></p></txBody></${name}:wrapper>`,
				{ txBody: { p: paragraph } },
			);
			expect(smartArtChildOrder(item)).toStrictEqual(['rPr']);
		},
	);

	it.each(['', 'a:', 'drawing_1:'])(
		'keeps attributed and multiline paragraphs with prefix %j',
		(prefix) => {
			const paragraph: XmlObject = {
				[`${prefix}r`]: { [`${prefix}t`]: 'First' },
				[`${prefix}tab`]: '',
			};
			annotateSmartArtTextOrder(
				`<${prefix}txBody><${prefix}p\n id="1"><${prefix}r><${prefix}t>First</${prefix}t></${prefix}r><${prefix}tab/></${prefix}p></${prefix}txBody>`,
				{ [`${prefix}txBody`]: { [`${prefix}p`]: paragraph } },
			);
			expect(smartArtChildOrder(paragraph)).toStrictEqual(['r', 'tab']);
		},
	);

	it.each([null, undefined, '', 0, false])('ignores primitive parser result %j', (parsed) => {
		expect(() => annotateSmartArtTextOrder('<a:txBody><a:p/></a:txBody>', parsed)).not.toThrow();
	});

	it('filters primitive paragraph array values before WeakMap annotation', () => {
		const paragraph: XmlObject = {
			'a:r': [{ 'a:t': 'First' }, { 'a:t': 'Second' }],
			'a:tab': '',
		};
		const parsed = {
			'a:txBody': {
				'a:p': [null, '', 7, paragraph],
			},
		};

		expect(() =>
			annotateSmartArtTextOrder(
				'<a:txBody><a:p><a:r><a:t>First</a:t></a:r><a:tab/><a:r><a:t>Second</a:t></a:r></a:p></a:txBody>',
				parsed,
			),
		).not.toThrow();
		expect(smartArtChildOrder(paragraph)).toStrictEqual(['r', 'tab', 'r']);
		expect(orderedSmartArtTextEntries(paragraph).map(([key]) => key)).toStrictEqual([
			'a:r',
			'a:tab',
			'a:r',
		]);
	});
});

/** Best of three runs, in milliseconds. */

describe('smartArt text order scan', () => {
	it('is linear in empty breaks and paragraphs that never close', () => {
		const parsed = { 'dgm:dataModel': { 'dgm:t': { 'a:p': { 'a:r': {} } } } };
		const run = (count: number) => () => {
			const xml = `<dgm:dataModel><dgm:t><a:p>${'<a:br/>'.repeat(count)}${'<a:r>'.repeat(count)}</a:p></dgm:t></dgm:dataModel>`;
			annotateSmartArtTextOrder(xml, {});
			annotateSmartArtTextOrder(xml, parsed);
		};
		expect(growth(run(20_000), run(80_000))).toBeLessThan(8);
	});

	it('gives an empty element an entry only when it parses to an object', () => {
		const first: XmlObject = { r: { t: 'One' } };
		const second: XmlObject = { r: [{ t: 'Two' }, { t: 'Three' }], br: '' };
		annotateSmartArtTextOrder(
			'<dgm:t><a:p><a:r><a:t>One</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>Two</a:t></a:r><a:br/><a:r><a:t>Three</a:t></a:r></a:p></dgm:t>',
			{ t: { p: [first, '', second] } },
		);
		expect(smartArtChildOrder(first)).toStrictEqual(['r']);
		expect(smartArtChildOrder(second)).toStrictEqual(['r', 'br', 'r']);
	});
});
