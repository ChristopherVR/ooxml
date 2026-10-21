import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel } from './index';
import { expectParagraph } from './test-support/access';
import {
	complexField,
	paragraph,
	simpleField,
	styleRefFixture,
} from './test-support/style-ref-fixture';

const HEADER = paragraph(
	'',
	undefined,
	`<w:r><w:t xml:space="preserve">Document </w:t></w:r>${complexField('STYLEREF "Doc Number" \\* MERGEFORMAT', ['DOC', '-1'])}`,
);
const FOOTER = paragraph('', undefined, simpleField('STYLEREF Title', 'Plan'));

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const found = new Map<string, string>();
	for (const file of Object.values(zip.files))
		if (!file.dir) found.set(file.name, await file.async('string'));
	return found;
}

/** Sets the text of the body paragraph in `style`. */
function setText(model: DocumentModel, style: string, text: string): DocumentModel {
	const next = structuredClone(model);
	const target = next.blocks.find((block) => block.type === 'paragraph' && block.style === style);
	expectParagraph(target).runs = [{ text }];
	return next;
}

describe('saving STYLEREF fields', () => {
	it('refreshes header and footer results from the edited body and leaves other parts alone', async () => {
		const original = await styleRefFixture({
			body: paragraph('Plan', 'Title') + paragraph('DOC-1', 'DocNumber') + paragraph('Notes'),
			header: HEADER,
			footer: FOOTER,
		});
		const loaded = await loadDocx(original);
		const saved = await loaded.save(
			setText(setText(loaded.model, 'DocNumber', 'DOC-0042'), 'Title', 'Survey'),
		);
		const before = await parts(original);
		const after = await parts(saved);
		const header = after.get('word/header1.xml')!;
		expect(header).toContain('>Document <');
		expect(header).toMatch(/<w:b\/><\/w:rPr><w:t[^>]*>DOC-0042<\/w:t>/);
		expect(header).not.toContain('-1<');
		expect(after.get('word/footer1.xml')).toContain('>Survey<');
		for (const name of ['word/styles.xml', '[Content_Types].xml', 'word/_rels/document.xml.rels'])
			expect(after.get(name), name).toBe(before.get(name));
		// Reopened, the header still holds one field whose result is the new number.
		const reopened = (await loadDocx(saved)).model;
		const runs = expectParagraph(reopened.sections![0]!.headers!.default!.blocks[0]).runs;
		expect(
			runs
				.filter((run) => run.field)
				.map((run) => run.text)
				.join(''),
		).toBe('DOC-0042');
		expect(runs.filter((run) => run.fieldChar)).toHaveLength(3);
	});

	it('returns the original bytes for an unchanged document, even with a stale header result', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-0042', 'DocNumber'),
			header: HEADER,
		});
		const loaded = await loadDocx(original);
		expect(await loaded.save()).toEqual(original);
	});

	it('does not rewrite a header whose result is already current', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber') + paragraph('Notes'),
			header: paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1'])),
		});
		const loaded = await loadDocx(original);
		const model = structuredClone(loaded.model);
		expectParagraph(model.blocks[1]).runs = [{ text: 'More notes' }];
		const saved = await parts(await loaded.save(model));
		expect(saved.get('word/header1.xml')).toBe((await parts(original)).get('word/header1.xml'));
	});

	it('saves without the refresh when the header cannot be rebuilt', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				// Bidirectional override text: inline markup the writer cannot rebuild.
				`${complexField('STYLEREF "Doc Number"', ['DOC-1'])}<w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>`,
			),
		});
		const loaded = await loadDocx(original);
		const saved = await parts(await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-2')));
		expect(saved.get('word/document.xml')).toContain('>DOC-2<');
		expect(saved.get('word/header1.xml')).toBe((await parts(original)).get('word/header1.xml'));
	});
});
