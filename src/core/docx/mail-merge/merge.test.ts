import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { documentParagraphs } from '../cell-blocks';
import { loadDocx } from '../parse';
import type { Block } from '../model';
import {
	checkMailMerge,
	finishMailMerge,
	finishMailMergeIndividually,
	mailMergeDelimitedText,
} from './index';
import { code, complex, paragraph, run, simple, template } from './test-fixture';

const text = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((paragraph) => paragraph.runs.map((run) => run.text).join(''));

describe('Word mail merge document output', () => {
	it('reopens combined sections and individual documents with static field values and personalized headers', async () => {
		const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>
			${paragraph(complex(code(' NEXT ')) + simple('MERGEFIELD Name') + run(' ') + simple('MERGEREC'))}
		</w:tc></w:tr></w:tbl>`;
		const loaded = await template(
			paragraph(simple('MERGEFIELD FirstName \\m') + run(' ') + simple('MERGESEQ')) +
				table +
				'<w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>',
			{
				'word/header1.xml': `<w:hdr xmlns:w="${NS.w}">${paragraph(run('Dear ') + simple('MERGEFIELD Name'))}</w:hdr>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/header" Target="header1.xml"/></Relationships>`,
				'word/settings.xml': `<w:settings xmlns:w="${NS.w}"><w:mailMerge><w:odso><w:fieldMapData><w:name w:val="Name"/><w:mappedName w:val="FirstName"/></w:fieldMapData></w:odso></w:mailMerge></w:settings>`,
			},
		);
		const data = await mailMergeDelimitedText('Name\r\nAda\r\nBea\r\nCal').read();
		const merged = finishMailMerge(loaded.model, data);
		expect(merged.issues).toEqual([]);
		expect(merged.mergedCount).toBe(2);
		const combined = await loadDocx(await loaded.save(merged.document));
		expect(text(combined.model.blocks)).toEqual(['Ada 1', 'Bea 2', '', 'Cal 2', ' ']);
		expect(combined.model.sections).toHaveLength(2);
		expect(combined.model.sections?.[1]?.type).toBe('nextPage');
		expect(
			combined.model.sections?.map((section) => text(section.headers?.default?.blocks ?? [])),
		).toEqual([['Dear Ada'], ['Dear Cal']]);
		expect(combined.model.mailMerge).toBeUndefined();
		for (const p of documentParagraphs(combined.model.blocks))
			expect(
				p.runs.every((run) => !run.field && !run.fieldChar && run.fieldCode === undefined),
			).toBe(true);
		const individual = finishMailMergeIndividually(loaded.model, data);
		expect(individual.documents.map((item) => item.recordNumber)).toEqual([1, 3]);
		const reopened = await Promise.all(
			individual.documents.map(async ({ document }) => loadDocx(await loaded.save(document))),
		);
		expect(reopened.map(({ model }) => text(model.blocks))).toEqual([
			['Ada 1', 'Bea 2'],
			['Cal 2', ' '],
		]);
		expect(await loaded.save()).toEqual(loaded.bytes);
	});

	it('returns the same issues from check and finish, with no document when every record is skipped', async () => {
		const loaded = await template(paragraph(simple('MERGEFIELD Absent') + simple('SKIPIF 1 = 1')));
		const data = { records: [{ Name: 'Ada' }] };
		const result = finishMailMerge(loaded.model, data);
		expect(result).toEqual({ mergedCount: 0, issues: checkMailMerge(loaded.model, data) });
		expect(result.issues).toEqual([
			{ code: 'unknown-field', recordNumber: 1, field: 'Absent', blockId: 'p0' },
		]);
	});

	it('preserves nested tables and source-only properties in repeated blocks, and gives notes independent ids', async () => {
		const table = (content: string) =>
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${content}</w:tc></w:tr></w:tbl>`;
		const source = `<w:p><w:pPr><w:suppressAutoHyphens/></w:pPr>${simple('MERGEFIELD Name')}</w:p>`;
		const loaded = await template(
			table(paragraph(run('before')) + table(source) + paragraph(run('after'))) +
				paragraph('<w:r><w:footnoteReference w:id="1"/></w:r>'),
			{
				'word/footnotes.xml': `<w:footnotes xmlns:w="${NS.w}"><w:footnote w:id="1">${paragraph(simple('MERGEFIELD Name'))}</w:footnote></w:footnotes>`,
			},
		);
		const result = finishMailMerge(loaded.model, { records: [{ Name: 'Ada' }, { Name: 'Bea' }] });
		expect(result.issues).toEqual([]);
		const bytes = await loaded.save(result.document);
		const reopened = await loadDocx(bytes);
		expect(text(reopened.model.blocks)).toEqual([
			'before',
			'Ada',
			'after',
			'',
			'before',
			'Bea',
			'after',
			'',
		]);
		expect(reopened.model.footnotes?.map((note) => [note.id, text(note.blocks)])).toEqual([
			['1', ['Ada']],
			['2', ['Bea']],
		]);
		expect(
			documentParagraphs(reopened.model.blocks).flatMap((p) =>
				p.runs.flatMap((run) => run.noteReference?.id ?? []),
			),
		).toEqual(['1', '2']);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string');
		expect(xml?.match(/<w:suppressAutoHyphens\s*\/>/g)).toHaveLength(2);
	});

	it('starts each copy on a new page without carrying over a later section header', async () => {
		const loaded = await template(
			'<w:p><w:pPr><w:sectPr/></w:pPr></w:p>' +
				paragraph(simple('MERGEFIELD Name')) +
				'<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:type w:val="continuous"/></w:sectPr>',
			{
				'word/header1.xml': `<w:hdr xmlns:w="${NS.w}">${paragraph(simple('MERGEFIELD Name'))}</w:hdr>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/header" Target="header1.xml"/></Relationships>`,
			},
		);
		const result = finishMailMerge(loaded.model, { records: [{ Name: 'Ada' }, { Name: 'Bea' }] });
		const reopened = await loadDocx(await loaded.save(result.document));
		expect(reopened.model.sections?.map((section) => section.type)).toEqual([
			'nextPage',
			'continuous',
			'nextPage',
			'continuous',
		]);
		expect(
			reopened.model.sections?.map(
				(section) => section.headers?.default?.blocks && text(section.headers.default.blocks),
			),
		).toEqual([undefined, ['Ada'], [''], ['Bea']]);
	});
});
