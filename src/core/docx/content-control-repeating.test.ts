import { describe, expect, it } from 'vitest';
import { buildXml, parseXml, WORD_NS } from './xml';
import { parseContentControl, WORD_2012_NS } from './content-control';
import { createContentControl } from './write-content-control';
import type { ContentControlLevel } from './content-control-model';

const create = (kind: 'repeatingSection' | 'repeatingSectionItem', level: ContentControlLevel) => {
	const doc = parseXml(`<w:document xmlns:w="${WORD_NS}"><w:body/></w:document>`);
	const { sdt } = createContentControl(doc, { id: '7', level, kind, alias: 'Fees' });
	doc.documentElement.firstChild!.appendChild(sdt);
	return { doc, sdt };
};

describe('repeating section content controls', () => {
	it.each(['repeatingSection', 'repeatingSectionItem'] as const)(
		'writes and reads back a block or row %s',
		(kind) => {
			for (const level of ['block', 'row'] as const) {
				const { doc, sdt } = create(kind, level);
				const xml = buildXml(doc);
				expect(xml).toContain(`<w15:${kind}/></w:sdtPr>`);
				expect(xml).toContain(`xmlns:w15="${WORD_2012_NS}"`);
				expect(xml).toMatch(/mc:Ignorable="[^"]*w15/);
				expect(parseContentControl(sdt, level)).toMatchObject({ kind, alias: 'Fees' });
			}
		},
	);

	it('writes no repeating type for an inline or cell control', () => {
		for (const level of ['inline', 'cell'] as const) {
			const xml = buildXml(create('repeatingSection', level).doc);
			expect(xml).not.toContain('repeatingSection');
		}
	});

	it('reads an item that Word wrote as its own kind', () => {
		const sdt = parseXml(
			`<w:sdt xmlns:w="${WORD_NS}" xmlns:w15="${WORD_2012_NS}"><w:sdtPr><w:id w:val="3"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent/></w:sdt>`,
		).documentElement;
		expect(parseContentControl(sdt, 'row').kind).toBe('repeatingSectionItem');
	});
});
