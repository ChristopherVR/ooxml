// Optional controls, locks, flattening, reading nested controls back, headers and footers, nested
// tables, what is kept, and filling an already filled document again from its template.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { documentParagraphs } from '../cell-blocks';
import type { Block, DocumentModel } from '../model';
import { listFormFields } from '../form-fill-fields';
import { fillFormFields } from '../form-fill-write';
import { loadDocx } from '../parse';
import { fillContentControls, flattenContentControls, readContentControls } from './index';
import {
	cell,
	item,
	paragraph,
	run,
	saveAndReload,
	section,
	sdt,
	table,
	template,
} from './test-fixture';

const texts = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((entry) => entry.runs.map((r) => r.text).join(''));
const SUFFIX = { optionalSuffix: '_optional' };

describe('optional content controls', () => {
	const optional = () =>
		template(
			paragraph(run('Start')) +
				sdt(
					'notes_optional',
					paragraph(run('Notes: ') + sdt('notes', run('n'), '<w:showingPlcHdr/>')),
				) +
				sdt('extra_optional', paragraph(run('Static text'))) +
				paragraph(run('A') + sdt('inline_optional', run('i'), '<w:showingPlcHdr/>') + run('B')) +
				table(
					`<w:tr>${cell(paragraph(run('h')))}${cell(paragraph(run('h')))}</w:tr>` +
						`<w:tr>${cell(paragraph(sdt('row_optional', run('r'))))}${cell(paragraph(run('')))}</w:tr>`,
				) +
				paragraph(run('End')),
		);

	it('removes an optional control when it and every control inside it are empty', async () => {
		const loaded = await optional();
		const filled = fillContentControls(
			loaded.model,
			{ values: { notes: '', row_optional: ' ' } },
			SUFFIX,
		);
		expect(
			filled.results.filter((result) => result.status === 'removed').map((result) => result.tag),
		).toEqual(['notes_optional', 'extra_optional', 'inline_optional', 'row_optional']);
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual(['Start', 'AB', 'h', 'h', 'End']);
		expect(reopened.xml).not.toContain('<w:sdt>');
		expect(reopened.xml.match(/<w:tr>/g)).toHaveLength(1);
	});

	it('keeps an optional control when a control inside it has a value', async () => {
		const loaded = await optional();
		const filled = fillContentControls(
			loaded.model,
			{ values: { notes: 'hello', inline_optional: 'x' } },
			SUFFIX,
		);
		const reopened = await saveAndReload(loaded, filled.model);
		// The row's optional control keeps the text it has.
		expect(texts(reopened.model.blocks)).toEqual([
			'Start',
			'Notes: hello',
			'AxB',
			'h',
			'h',
			'r',
			'',
			'End',
		]);
		expect(filled.results).toContainEqual({
			tag: 'extra_optional',
			status: 'removed',
			story: 'body',
		});
		expect(filled.results).toContainEqual({ tag: 'notes_optional', status: 'kept', story: 'body' });
	});
});

describe('locks, reading back and flattening', () => {
	it('sets locks by tag after filling, on nested and repeated controls', async () => {
		const loaded = await template(
			sdt(
				'group',
				paragraph(
					sdt('name', run('n')) + run(' ') + sdt('free', run('f'), '<w:lock w:val="sdtLocked"/>'),
				),
			),
		);
		const filled = fillContentControls(
			loaded.model,
			{ values: { name: 'Ada' } },
			{ lock: { group: 'sdtLocked', name: 'sdtContentLocked', free: 'none' } },
		);
		const reopened = await saveAndReload(loaded, filled.model);
		const xml = reopened.xml;
		expect(xml.match(/<w:lock w:val="sdtLocked"\/>/g)).toHaveLength(1);
		expect(xml.match(/<w:lock w:val="sdtContentLocked"\/>/g)).toHaveLength(1);
		const runs = documentParagraphs(reopened.model.blocks)[0]!.runs;
		expect(runs.find((r) => r.text === 'Ada')?.contentControl?.lock).toBe('sdtContentLocked');
		expect(runs.find((r) => r.text === 'f')?.contentControl?.lock).toBeUndefined();
		expect(reopened.model.blocks[0]?.contentControl?.lock).toBe('sdtLocked');
	});

	it('reads wrappers through to the controls inside them, table cells included', async () => {
		const loaded = await template(
			sdt(
				'outer',
				paragraph(run('Hi ') + sdt('inner', run('Ada'))) + sdt('deep', paragraph(run('Deep'))),
			) +
				table(
					`<w:tr>${cell(paragraph(sdt('cellText', run('C'))))}${cell(paragraph(run('x')))}</w:tr>`,
				) +
				paragraph(sdt('empty', run('Click here'), '<w:showingPlcHdr/>')),
		);
		expect(readContentControls(loaded.model)).toEqual({
			outer: 'Hi Ada\nDeep',
			inner: 'Ada',
			deep: 'Deep',
			cellText: 'C',
			empty: '',
		});
		expect(readContentControls(loaded.model, { prefix: 'de' })).toEqual({ deep: 'Deep' });
		expect(readContentControls(loaded.model, { as: 'html' }).outer).toBe(
			'<p>Hi Ada</p><p>Deep</p>',
		);
	});

	it('round trips: reading a filled document gives the values filled', async () => {
		const loaded = await template(
			sdt('rich', paragraph(run('r'))) +
				paragraph(sdt('plain', run('p'))) +
				table(`<w:tr>${cell(paragraph(sdt('cellV', run('c'))))}</w:tr>`),
		);
		const values = { rich: '<p><b>Bold</b> and <i>it</i></p>', plain: 'Just text', cellV: 'Cell' };
		const filled = fillContentControls(loaded.model, { values, html: ['rich'] });
		const reopened = await saveAndReload(loaded, filled.model);
		expect(readContentControls(reopened.model, { as: 'html' })).toEqual({
			rich: '<p><b>Bold</b> and <i>it</i></p>',
			plain: 'Just text',
			cellV: 'Cell',
		});
		expect(readContentControls(reopened.model)).toEqual({
			rich: 'Bold and it',
			plain: 'Just text',
			cellV: 'Cell',
		});
	});

	it('flattens every control into its content and drops hidden paragraphs', async () => {
		const loaded = await template(
			sdt('a', paragraph(run('Keep ') + sdt('b', run('this')))) +
				'<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>' +
				table(sdt('row', `<w:tr>${cell(paragraph(sdt('c', run('cell'))))}</w:tr>`)) +
				paragraph(run('end')),
		);
		const flat = flattenContentControls(loaded.model, { removeHiddenParagraphs: true });
		const reopened = await saveAndReload(loaded, flat);
		expect(reopened.xml).not.toContain('<w:sdt');
		expect(texts(reopened.model.blocks)).toEqual(['Keep this', 'cell', 'end']);
		const kept = await saveAndReload(loaded, flattenContentControls(loaded.model));
		expect(texts(kept.model.blocks)).toContain('hidden');
	});
});

const rels = (...items: [id: string, type: string, target: string][]) =>
	`<Relationships xmlns="${NS.rels}">${items
		.map(
			([id, type, target]) =>
				`<Relationship Id="${id}" Type="${NS.r}/${type}" Target="${target}"/>`,
		)
		.join('')}</Relationships>`;

describe('stories, nested tables and what is kept', () => {
	it('fills controls in headers and footers', async () => {
		const loaded = await template(
			paragraph(sdt('body', run('b'))) +
				'<w:sectPr><w:headerReference w:type="default" r:id="rH"/><w:footerReference w:type="default" r:id="rF"/></w:sectPr>',
			{
				'word/header1.xml': `<w:hdr xmlns:w="${NS.w}">${paragraph(sdt('head', run('h')))}</w:hdr>`,
				'word/footer1.xml': `<w:ftr xmlns:w="${NS.w}">${sdt('foot', paragraph(run('f')))}</w:ftr>`,
				'word/_rels/document.xml.rels': rels(
					['rH', 'header', 'header1.xml'],
					['rF', 'footer', 'footer1.xml'],
				),
			},
		);
		const filled = fillContentControls(loaded.model, {
			values: { body: 'B', head: 'H', foot: 'F' },
		});
		expect(filled.results.map((result) => [result.tag, result.story, result.slot])).toEqual([
			['body', 'body', undefined],
			['head', 'header', 'default'],
			['foot', 'footer', 'default'],
		]);
		const bytes = await loaded.save(filled.model);
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('word/header1.xml')!.async('string')).toContain('>H<');
		expect(await zip.file('word/footer1.xml')!.async('string')).toContain('>F<');
		expect(readContentControls((await loadDocx(bytes)).model)).toEqual({
			body: 'B',
			head: 'H',
			foot: 'F',
		});
	});

	it('fills a control in a table nested in a cell', async () => {
		const inner = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${paragraph(sdt('deep', run('D')))}</w:tc></w:tr></w:tbl>`;
		const loaded = await template(
			table(
				`<w:tr>${cell(paragraph(run('a')) + inner + paragraph(run('')))}${cell(paragraph(run('b')))}</w:tr>`,
			),
		);
		const filled = fillContentControls(loaded.model, { values: { deep: 'Nested' } });
		expect(filled.results).toEqual([{ tag: 'deep', status: 'filled', story: 'body' }]);
		const reopened = await saveAndReload(loaded, filled.model);
		expect(reopened.xml).toContain('>Nested<');
		expect(readContentControls(reopened.model)).toEqual({ deep: 'Nested' });
		// fillFormFields no longer refuses it either.
		const field = listFormFields(loaded.model).find((entry) => entry.tag === 'deep')!;
		expect(field).toMatchObject({ fillable: true });
		const form = fillFormFields(loaded.model, { [field.id]: 'By form' });
		expect(form.results).toEqual([{ id: field.id, ok: true, value: 'By form' }]);
		expect((await saveAndReload(loaded, form.model)).xml).toContain('>By form<');
	});

	it('keeps controls not named, bookmarks, comments and custom XML parts', async () => {
		const loaded = await template(
			`<w:p><w:commentRangeStart w:id="0"/>${run('Dear ')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${sdt('name', run('n'))}<w:bookmarkStart w:id="1" w:name="anchor"/>${run('.')}<w:bookmarkEnd w:id="1"/>${sdt('other', run('untouched'))}</w:p>`,
			{
				'word/comments.xml': `<w:comments xmlns:w="${NS.w}"><w:comment w:id="0" w:author="Rev">${paragraph(run('check'))}</w:comment></w:comments>`,
				'word/_rels/document.xml.rels': rels(['rC', 'comments', 'comments.xml']),
				'customXml/item1.xml': '<data xmlns="urn:example"><id>42</id></data>',
			},
		);
		const filled = fillContentControls(loaded.model, { values: { name: 'Ada' } });
		const { xml, bytes } = await saveAndReload(loaded, filled.model);
		expect(xml).toContain('<w:commentReference w:id="0"/>');
		expect(xml).toContain('w:name="anchor"');
		expect(xml).toContain('untouched');
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('customXml/item1.xml')!.async('string')).toContain('<id>42</id>');
		expect(await zip.file('word/comments.xml')!.async('string')).toContain('check');
	});
});

describe('filling an already filled document', () => {
	const build = () =>
		template(
			paragraph(run('Invoice ') + sdt('number', run('#'))) +
				sdt('note_optional', paragraph(sdt('note', run('n'), '<w:showingPlcHdr/>'))) +
				sdt(
					'items',
					sdt(undefined, paragraph(sdt('item_template_name', run('i'))), item),
					section,
				) +
				table(
					`<w:tr>${cell(paragraph(run('Line')))}</w:tr>` +
						`<w:tr>${cell(paragraph(sdt('line_template_text', run('l'))))}</w:tr>`,
				) +
				paragraph(run('Thanks')),
		);
	const fill = (
		model: DocumentModel,
		values: Record<string, string>,
		names: string[],
		lines: string[],
		template?: DocumentModel,
	) =>
		fillContentControls(
			model,
			{
				values,
				repeat: { items: names.map((name) => ({ name })), line: lines.map((text) => ({ text })) },
			},
			{ ...SUFFIX, ...(template && { template }) },
		).model;

	it('gives the result of filling the template, with the text typed outside the controls', async () => {
		const loaded = await build();
		const first = await saveAndReload(
			loaded,
			fill(loaded.model, { number: '1', note: '' }, ['A', 'B', 'C'], ['x', 'y']),
		);
		expect(texts(first.model.blocks)).toEqual([
			'Invoice 1',
			'A',
			'B',
			'C',
			'Line',
			'x',
			'y',
			'Thanks',
		]);
		// The user types a paragraph of their own after the first line.
		const edited: DocumentModel = {
			...first.model,
			blocks: [
				first.model.blocks[0]!,
				{ type: 'paragraph', id: 'mine', runs: [{ text: 'Typed by hand' }] },
				...first.model.blocks.slice(1),
			],
		};
		const second = loadDocx(await (await loadDocx(first.bytes)).save(edited));
		const reloaded = await second;
		const refilled = fill(
			reloaded.model,
			{ number: '2', note: 'Paid' },
			['D'],
			['z'],
			loaded.model,
		);
		const result = await saveAndReload(reloaded, refilled);
		expect(texts(result.model.blocks)).toEqual([
			'Invoice 2',
			'Typed by hand',
			'Paid',
			'D',
			'Line',
			'z',
			'Thanks',
		]);
		const fresh = await saveAndReload(
			loaded,
			fill(loaded.model, { number: '2', note: 'Paid' }, ['D'], ['z']),
		);
		expect(readContentControls(result.model)).toEqual(readContentControls(fresh.model));
		expect(result.xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(1);
	});
});
