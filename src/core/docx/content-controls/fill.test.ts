// Filling content controls with text and HTML, saved with the template's writer and read back.
import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import type { Block } from '../model';
import { fillContentControls, readContentControls } from './index';
import { paragraph, run, saveAndReload, sdt, styled, template } from './test-fixture';

const texts = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((item) => item.runs.map((r) => r.text).join(''));
const FONT = '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="28"/>';
const PLACEHOLDER = '<w:rStyle w:val="PlaceholderText"/>';

describe('fillContentControls', () => {
	it('fills inline and block controls with text in the control typography', async () => {
		const loaded = await template(
			paragraph(
				run('Dear ') + sdt('name', styled('Click here', FONT + PLACEHOLDER), '<w:showingPlcHdr/>'),
			) + sdt('notes', paragraph(styled('Notes here', FONT)) + paragraph(run('second'))),
		);
		const { model, results } = fillContentControls(loaded.model, {
			values: { name: 'Ada Lovelace', notes: 'Line one\nLine two' },
		});
		expect(results).toEqual([
			{ tag: 'name', status: 'filled', story: 'body' },
			{ tag: 'notes', status: 'filled', story: 'body' },
		]);
		const reopened = await saveAndReload(loaded, model);
		expect(texts(reopened.model.blocks)).toEqual(['Dear Ada Lovelace', 'Line one', 'Line two']);
		const name = documentParagraphs(reopened.model.blocks)[0]!.runs[1]!;
		expect(name).toMatchObject({ text: 'Ada Lovelace', fontFamily: 'Georgia', fontSize: 14 });
		expect(name.style).toBeUndefined();
		expect(name.contentControl).toMatchObject({ tag: 'name' });
		expect(name.contentControl?.showingPlaceholder).toBeUndefined();
		expect(reopened.xml).not.toContain('showingPlcHdr');
		// Both new paragraphs stay inside the block control.
		expect(reopened.xml.match(/<w:sdt>/g)).toHaveLength(2);
		const notes = reopened.model.blocks.slice(1);
		expect(notes.map((block) => block.contentControl?.tag)).toEqual(['notes', 'notes']);
		expect(readContentControls(reopened.model)).toEqual({
			name: 'Ada Lovelace',
			notes: 'Line one\nLine two',
		});
	});

	it('reads HTML into a block control, keeping its paragraph formatting and fonts', async () => {
		const loaded = await template(
			sdt(
				'body',
				'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
					styled('x', FONT) +
					'</w:p>',
			),
		);
		const html =
			'<p style="font-family:Arial;font-size:30px">One <b>bold</b> <span style="color:#ff0000">red</span></p>' +
			'<ul><li>first</li><li>second</li></ul><p>a<br>b <a href="https://example.com">link</a></p>' +
			'<table border="1"><tr><td>c1</td><td>c2</td></tr></table><script>alert(1)</script>';
		const filled = fillContentControls(loaded.model, { values: { body: html }, html: ['body'] });
		const reopened = await saveAndReload(loaded, filled.model);
		const blocks = reopened.model.blocks;
		expect(texts(blocks)).toEqual(['One bold red', 'first', 'second', 'a\nb link', 'c1', 'c2', '']);
		const paragraphs = documentParagraphs(blocks);
		for (const item of paragraphs.slice(0, 4)) {
			expect(item.align).toBe('center');
			expect(item.spacingAfterTwips).toBe(240);
			expect(item.contentControl?.tag).toBe('body');
			for (const r of item.runs.filter((each) => each.text.trim()))
				expect(r).toMatchObject({ fontFamily: 'Georgia', fontSize: 14 });
		}
		expect(paragraphs[0]!.runs.find((r) => r.text === 'bold')?.bold).toBe(true);
		expect(paragraphs[0]!.runs.find((r) => r.text === 'red')?.color?.toLowerCase()).toBe('#ff0000');
		expect(paragraphs[1]!.numbering).toBeDefined();
		expect(paragraphs[3]!.runs.find((r) => r.text === 'link')?.link?.href).toBe(
			'https://example.com',
		);
		expect(blocks.find((block) => block.type === 'table')?.contentControl?.tag).toBe('body');
		expect(reopened.xml).not.toContain('alert');
		expect(reopened.xml).not.toContain('Arial');
		const read = readContentControls(reopened.model, { as: 'html' });
		expect(read.body).toContain('<b>bold</b>');
		expect(read.body).toContain('<ul><li>first</li><li>second</li></ul>');
		expect(read.body).not.toMatch(/font/i);
		// Reading back and filling again gives the same content.
		const again = fillContentControls(reopened.model, {
			values: { body: read.body! },
			html: ['body'],
		});
		expect(texts(again.model.blocks)).toEqual(texts(blocks));
	});

	it('joins HTML paragraphs with line breaks in an inline control', async () => {
		const loaded = await template(paragraph(run('A: ') + sdt('a', run('x')) + run(' end')));
		const filled = fillContentControls(loaded.model, {
			values: { a: '<p>one</p><p><i>two</i></p>' },
			html: ['a'],
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual(['A: one\ntwo end']);
		expect(reopened.xml).toContain('<w:br/>');
		expect(readContentControls(reopened.model, { as: 'html' }).a).toBe('one<br><i>two</i>');
	});

	it('falls back to the text of invalid HTML', async () => {
		const loaded = await template(sdt('a', paragraph(run('x'))));
		const filled = fillContentControls(loaded.model, {
			values: { a: '<p>unclosed <b>bold <div><<>> & text' },
			html: ['a'],
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks).join(' ')).toContain('unclosed');
		expect(filled.results[0]).toMatchObject({ tag: 'a', status: 'filled' });
	});

	it.each([' ', '&nbsp;', '<p> </p>', '<p><b></b>&nbsp;</p>'])(
		'clears a control for the empty value %j, without its placeholder',
		async (value) => {
			const loaded = await template(
				sdt(
					'block',
					paragraph(styled('Type here', PLACEHOLDER)) + paragraph(run('more')),
					'<w:showingPlcHdr/>',
				) + paragraph(sdt('inline', styled('Type', PLACEHOLDER), '<w:showingPlcHdr/>')),
			);
			const filled = fillContentControls(loaded.model, {
				values: { block: value, inline: value },
				html: ['block', 'inline'],
			});
			expect(filled.results.map((result) => result.status)).toEqual(['cleared', 'cleared']);
			const reopened = await saveAndReload(loaded, filled.model);
			expect(texts(reopened.model.blocks)).toEqual(['', '']);
			expect(reopened.xml).not.toContain('Type');
			expect(reopened.xml).not.toContain('showingPlcHdr');
			expect(reopened.model.blocks[0]?.contentControl?.tag).toBe('block');
		},
	);

	it('sets check boxes, lists and dates, and says why a value does not fit', async () => {
		const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
		const loaded = await template(
			paragraph(
				sdt(
					'box',
					run('☐'),
					`<w14:checkbox xmlns:w14="${W14}"><w14:checked w14:val="0"/></w14:checkbox>`,
				) +
					sdt(
						'pick',
						run('Choose'),
						'<w:dropDownList><w:listItem w:displayText="Red" w:value="r"/><w:listItem w:displayText="Blue" w:value="b"/></w:dropDownList>',
					) +
					sdt(
						'day',
						run('date'),
						'<w:date><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/></w:date>',
					) +
					sdt(
						'bad',
						run('Choose'),
						'<w:dropDownList><w:listItem w:displayText="Red" w:value="r"/></w:dropDownList>',
					),
			),
		);
		const filled = fillContentControls(loaded.model, {
			values: { box: 'true', pick: 'b', day: '2026-03-04', bad: 'Green' },
		});
		expect(filled.results.find((result) => result.tag === 'bad')).toMatchObject({
			status: 'skipped',
			reason: expect.stringContaining('none of'),
		});
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual(['☒Blue2026-03-04Choose']);
		expect(readContentControls(reopened.model)).toMatchObject({
			box: 'true',
			pick: 'b',
			day: '2026-03-04',
		});
	});

	it('keeps data-bound and locked controls, and fills a locked one named in lock', async () => {
		const loaded = await template(
			paragraph(
				sdt('bound', run('B'), '<w:dataBinding w:xpath="/a" w:storeItemID="{1}"/>') +
					sdt('locked', run('L'), '<w:lock w:val="contentLocked"/>') +
					sdt('relock', run('R'), '<w:lock w:val="sdtContentLocked"/>'),
			),
		);
		const filled = fillContentControls(
			loaded.model,
			{ values: { bound: 'x', locked: 'y', relock: 'z' } },
			{ lock: { relock: 'sdtContentLocked' } },
		);
		expect(filled.results.map((result) => [result.tag, result.status])).toEqual([
			['bound', 'skipped'],
			['locked', 'skipped'],
			['relock', 'filled'],
			['relock', 'locked'],
		]);
		const reopened = await saveAndReload(loaded, filled.model);
		expect(texts(reopened.model.blocks)).toEqual(['BLz']);
		expect(reopened.xml).toContain('w:dataBinding');
	});

	it('reports the tags no control has', () => {
		return template(paragraph(run('x'))).then((loaded) => {
			const filled = fillContentControls(loaded.model, {
				values: { nope: 'x' },
				repeat: { rows: [] },
			});
			expect(filled.results).toEqual([
				{ tag: 'nope', status: 'missing', reason: 'no content control has the tag "nope"' },
				{ tag: 'rows', status: 'missing', reason: 'no content control has the tag "rows"' },
			]);
		});
	});
});
