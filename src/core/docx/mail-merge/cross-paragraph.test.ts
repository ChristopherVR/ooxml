import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import type { Block, Paragraph } from '../model';
import { loadDocx } from '../parse';
import { WORD_NS } from '../xml';
import { joinParagraphs, markSpaces, splitParagraphs } from './cross-paragraph';
import { finishMailMerge, mailMergeDelimitedText, previewMailMerge } from './index';
import { code, complex, marker, paragraph, run, simple, template } from './test-fixture';

const text = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((item) => item.runs.map((r) => r.text).join(''));
const records = (csv: string) => mailMergeDelimitedText(csv).read();

describe('merge fields across paragraphs', () => {
	it('keeps the paragraphs of an IF branch that holds paragraph marks', async () => {
		// { IF { MERGEFIELD Member } = "yes" "Welcome back.¶Your card is ready." "Join today." }
		const loaded = await template(
			paragraph(
				marker('begin') +
					code(' IF ') +
					complex(code(' MERGEFIELD Member '), 'yes') +
					code(' = "yes" "Welcome back.'),
			) +
				paragraph(
					code('Your card is ready." "Join today." ') +
						marker('separate') +
						run('saved') +
						marker('end') +
						run(' End'),
				) +
				paragraph(run('After')),
		);
		const data = await records('Member\nyes\nno');
		const first = previewMailMerge(loaded.model, data, 1);
		expect(first.issues).toEqual([]);
		expect(text(first.document.blocks)).toEqual([
			'Welcome back.',
			'Your card is ready. End',
			'After',
		]);
		const second = previewMailMerge(loaded.model, data, 2);
		expect(text(second.document.blocks)).toEqual(['Join today. End', 'After']);
		// The joined paragraph takes the properties and id of the paragraph whose mark ends it.
		const ids = loaded.model.blocks.map((block) => block.id);
		expect(second.document.blocks.map((block) => block.id)).toEqual([ids[1], ids[2]]);
		const merged = finishMailMerge(loaded.model, data);
		const reopened = await loadDocx(await loaded.save(merged.document));
		expect(text(reopened.model.blocks)).toEqual([
			'Welcome back.',
			'Your card is ready. End',
			'After',
			'Join today. End',
			'After',
		]);
	});

	it('replaces a merge field whose saved result spans paragraphs', async () => {
		const loaded = await template(
			paragraph(
				run('Dear ') +
					marker('begin') +
					code(' MERGEFIELD Name ') +
					marker('separate') +
					run('old'),
			) + paragraph(run('result') + marker('end') + run(',')),
		);
		const preview = previewMailMerge(loaded.model, await records('Name\nAda'), 1);
		expect(text(preview.document.blocks)).toEqual(['Dear Ada,']);
	});

	it('keeps a field that crosses a table or never closes', async () => {
		const table = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${paragraph(run('cell'))}</w:tc></w:tr></w:tbl>`;
		const loaded = await template(
			paragraph(marker('begin') + code(' QUOTE "a') + marker('separate') + run('kept')) +
				table +
				paragraph(marker('end') + simple('MERGEFIELD Name')),
		);
		const preview = previewMailMerge(loaded.model, await records('Name\nAda'), 1);
		expect(text(preview.document.blocks)).toEqual(['kept', 'cell', 'Ada']);
	});

	it('joins and splits paragraphs at their marks', () => {
		const group: Paragraph[] = [
			{ type: 'paragraph', id: 'a', runs: [{ text: 'one' }], alignment: 'center' },
			{ type: 'paragraph', id: 'b', runs: [{ text: 'two' }] },
		];
		const joined = joinParagraphs(group);
		expect(joined.runs.map((r) => r.text)).toEqual(['one', '', 'two']);
		expect(splitParagraphs(joined, group)).toEqual(group);
		const collapsed = { ...joined, runs: [{ text: 'onetwo' }] };
		expect(splitParagraphs(collapsed, group)).toEqual([
			{ type: 'paragraph', id: 'b', runs: [{ text: 'onetwo' }] },
		]);
		expect(markSpaces('IFa = "xy"')).toBe('IF a = "xy"');
	});
});

const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const textBox = (inner: string) =>
	`<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0"><wp:extent cx="1905000" cy="952500"/><wp:docPr id="7" name="Text Box 7"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${WPS}"><wps:wsp xmlns:wps="${WPS}"><wps:spPr/><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><w:t>stale</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

describe('merge fields in text boxes', () => {
	it('merges fields inside a DrawingML text box and saves the result into its text', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}"><w:body>${paragraph(
				textBox(paragraph(run('To ') + simple('MERGEFIELD Name')) + paragraph(run('Static'))) +
					run('Body ') +
					simple('MERGEFIELD Name'),
			)}<w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const data = await records('Name\nAda\nBea');
		const preview = previewMailMerge(loaded.model, data, 2);
		const image = (preview.document.blocks[0] as Paragraph).runs[0]!.image!;
		expect(text(image.textBoxContent!.blocks)).toEqual(['To Bea', 'Static']);
		const merged = finishMailMerge(loaded.model, data);
		const bytes = await loaded.save(merged.document);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('To </w:t>');
		expect(xml).toContain('>Ada<');
		expect(xml).toContain('>Bea<');
		expect(xml).not.toContain('MERGEFIELD');
		expect(xml).not.toContain('stale');
		const reopened = await loadDocx(bytes);
		const boxes = reopened.model.blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs.flatMap((r) => (r.image ? [r.image] : [])) : [],
		);
		expect(boxes.map((box) => text(box.textBoxContent!.blocks))).toEqual([
			['To Ada', 'Static'],
			['To Bea', 'Static'],
		]);
	});

	it('leaves a text box without merge fields byte for byte', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}"><w:body>${paragraph(
				textBox(paragraph(run('Static'))) + simple('MERGEFIELD Name'),
			)}<w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const merged = finishMailMerge(loaded.model, await records('Name\nAda'));
		const xml = await (
			await JSZip.loadAsync(await loaded.save(merged.document))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('stale');
	});
});
