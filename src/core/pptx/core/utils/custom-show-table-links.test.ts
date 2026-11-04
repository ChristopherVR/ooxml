import { describe, expect, it } from 'vitest';
import type { XmlObject } from '../types';
import { mapCustomShowTableLinks } from './custom-show-table-links';
import { xmlChildren, xmlPath } from './xml-access';

function link(id: string): XmlObject {
	return {
		'@_action': 'ppaction://customshow?id=' + id + '&return=true',
		'@_r:id': '',
		'@_tooltip': 'tip',
		'a:snd': { '@_r:embed': 'sound' },
	};
}
function source() {
	return {
		'p:nvGraphicFramePr': { 'p:cNvPr': { 'a:hlinkClick': link('7') } },
		'a:graphic': {
			'a:graphicData': {
				'a:tbl': {
					'a:tr': [
						{
							'a:tc': [
								{
									'a:txBody': {
										'a:p': {
											'a:r': [
												{
													'a:rPr': {
														'@_b': '1',
														'a:hlinkClick': link('7'),
														'a:hlinkMouseOver': link('8'),
													},
													'a:t': 'First',
												},
												{ 'a:rPr': { 'a:hlinkClick': link('8') }, 'a:t': 'Second' },
											],
											'a:br': { 'a:rPr': { '@_lang': 'de-DE' } },
											'a:fld': {
												'@_id': 'field',
												'a:rPr': { 'a:hlinkClick': link('7') },
												'a:t': 'Field',
											},
											'a:endParaRPr': { '@_lang': 'en-US' },
										},
									},
								},
							],
						},
					],
				},
			},
		},
	};
}
function paragraph(raw: XmlObject): XmlObject {
	const table = xmlPath(raw, 'a:graphic', 'a:graphicData', 'a:tbl')!;
	const row = xmlChildren(table, 'a:tr')[0]!;
	const cell = xmlChildren(row, 'a:tc')[0]!;
	return xmlPath(cell, 'a:txBody', 'a:p')!;
}
function freeze(value: unknown): void {
	if (!value || typeof value !== 'object') return;
	for (const item of Object.values(value)) freeze(item);
	Object.freeze(value);
}

describe('preserved custom-show table links', () => {
	it('clears matching rich-run and field links, preserving unrelated links and source objects', () => {
		const raw = source();
		const before = structuredClone(raw);
		freeze(raw);
		const mapped = mapCustomShowTableLinks(raw, (id) => (id === '7' ? null : id));
		expect(mapped.changed).toBe(2);
		expect(raw).toStrictEqual(before);
		const oldParagraph = paragraph(raw);
		const next = paragraph(mapped.rawXml);
		const oldRuns = xmlChildren(oldParagraph, 'a:r');
		const runs = xmlChildren(next, 'a:r');
		expect(xmlPath(runs[0], 'a:rPr', 'a:hlinkClick')).toBeUndefined();
		expect(xmlPath(next, 'a:fld', 'a:rPr', 'a:hlinkClick')).toBeUndefined();
		expect(xmlPath(runs[0], 'a:rPr', 'a:hlinkMouseOver')).toBe(
			xmlPath(oldRuns[0], 'a:rPr', 'a:hlinkMouseOver'),
		);
		expect(runs[1]).toBe(oldRuns[1]);
		expect(next['a:br']).toBe(oldParagraph['a:br']);
		expect(next['a:endParaRPr']).toBe(oldParagraph['a:endParaRPr']);
		expect(mapped.rawXml['p:nvGraphicFramePr']).toBe(raw['p:nvGraphicFramePr']);
	});

	it('rewrites only the show id, preserving the target mode, tooltip and sound', () => {
		const raw = source();
		const mapped = mapCustomShowTableLinks(raw, (id) => (id === '7' ? '12' : id));
		const run = xmlChildren(paragraph(mapped.rawXml), 'a:r')[0]!;
		expect(xmlPath(run, 'a:rPr', 'a:hlinkClick')).toEqual({
			...link('7'),
			'@_action': 'ppaction://customshow?id=12&return=true',
		});
		expect(mapped.changed).toBe(2);
		const unchanged = mapCustomShowTableLinks(raw, (id) => id);
		expect(unchanged.rawXml).toBe(raw);
		expect(unchanged.changed).toBe(0);
	});
});
