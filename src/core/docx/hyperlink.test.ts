import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx } from './index';
import { isWordNavigableHref } from './hyperlink';
import { expectParagraph } from './test-support/access';

describe('isWordNavigableHref', () => {
	it('accepts the schemes Word follows: http, https, mailto, tel and ftp', () => {
		for (const href of [
			'http://a',
			'HTTPS://a',
			' mailto:x@y.z',
			'tel:+27215550100',
			'TEL:555',
			'ftp://files.example.com/a.txt',
		])
			expect(isWordNavigableHref(href)).toBe(true);
	});

	it('rejects script, disguised, local and unknown targets', () => {
		for (const href of [
			'javascript:alert(1)',
			'java\u0000script:alert(1)',
			'data:text/html,x',
			'file:///c:/x',
			'C:\\Reports\\q1.docx',
			'/relative',
			'sms:555',
			'',
			undefined,
			42,
		])
			expect(isWordNavigableHref(href)).toBe(false);
	});
});

describe('tel and ftp hyperlinks', () => {
	it('round-trip as external relationships', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{ text: 'Call', link: { href: 'tel:+27215550100' } },
					{ text: ' or ' },
					{ text: 'download', link: { href: 'ftp://files.example.com/a.txt' } },
				],
			},
		];
		const bytes = await saveDocx(model);
		const rels =
			(await (
				await JSZip.loadAsync(bytes)
			)
				.file('word/_rels/document.xml.rels')
				?.async('string')) ?? '';
		expect(rels).toContain('Target="tel:+27215550100" TargetMode="External"');
		expect(rels).toContain('Target="ftp://files.example.com/a.txt" TargetMode="External"');
		const reopened = expectParagraph((await loadDocx(bytes)).model.blocks[0]);
		expect(reopened.runs.map((run) => run.link?.href)).toEqual([
			'tel:+27215550100',
			undefined,
			'ftp://files.example.com/a.txt',
		]);
	});
});
