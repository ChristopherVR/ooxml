import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel, type Paragraph } from './index';
import { refreshStyleRefFields } from './style-ref-fields';
import { at, expectParagraph, must } from './test-support/access';
import {
	complexField,
	paragraph,
	simpleField,
	styleRefFixture,
} from './test-support/style-ref-fixture';

const load = async (parts: Parameters<typeof styleRefFixture>[0]) =>
	(await loadDocx(await styleRefFixture(parts))).model;
const headerParagraph = (model: DocumentModel): Paragraph =>
	expectParagraph(must(at(model.sections, 0).headers?.default, 'header').blocks[0]);
const shown = (item: Paragraph) => item.runs.map((run) => run.text).join('');

describe('refreshing STYLEREF results', () => {
	it('gives a header field the text it shows on every page, in its first result run', async () => {
		const model = await load({
			body: paragraph('DOC-0042', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				complexField('STYLEREF "Doc Number" \\* MERGEFORMAT', ['DOC', '-1']),
			),
			footer: paragraph('', undefined, simpleField('STYLEREF DocNumber', 'DOC-1')),
		});
		const refreshed = refreshStyleRefFields(model);
		const header = headerParagraph(refreshed);
		const results = header.runs.filter((run) => run.field);
		expect(results.map((run) => run.text)).toEqual(['DOC-0042', '']);
		expect(results[0]!.bold).toBe(true);
		const footer = expectParagraph(
			must(at(refreshed.sections, 0).footers?.default, 'footer').blocks[0],
		);
		expect(shown(footer)).toBe('DOC-0042');
		expect(refreshed.blocks).toBe(model.blocks);
	});

	it('returns the model itself when every result is current', async () => {
		const model = await load({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1'])),
		});
		expect(refreshStyleRefFields(model)).toBe(model);
	});

	it('keeps a header result that depends on the page, a locked field and an unknown style', async () => {
		const model = await load({
			body: paragraph('One', 'Heading1') + paragraph('Two', 'Heading1'),
			header: paragraph(
				'',
				undefined,
				complexField('STYLEREF 1', ['Cached']) +
					complexField(
						'STYLEREF "Doc Number"',
						['Locked'],
						'<w:fldChar w:fldCharType="begin" w:fldLock="1"/>',
					) +
					complexField('STYLEREF Missing', ['Kept']),
			),
		});
		const withNumber = {
			...model,
			blocks: [
				...model.blocks,
				{ type: 'paragraph' as const, id: 'doc', style: 'DocNumber', runs: [{ text: 'DOC-9' }] },
			],
		};
		expect(refreshStyleRefFields(withNumber).sections).toBe(withNumber.sections);
	});

});
