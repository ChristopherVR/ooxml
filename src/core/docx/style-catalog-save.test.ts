import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createDocument, editStyleCatalogs, loadDocx, saveDocx, type StyleEdit } from './index';
import { must } from './test-support/access';
import { twips } from './units';

const createStyles = [
	{
		kind: 'paragraph',
		name: 'Article Heading',
		basedOn: 'Normal',
		next: 'Normal',
		uiPriority: 7,
		quickFormat: false,
		paragraphFormatting: { align: 'center' },
		runFormatting: { fontSize: 16 },
	},
	{ kind: 'character', name: 'Gloss', runFormatting: { italic: true, fontFamily: 'Georgia' } },
	{
		kind: 'linked',
		name: 'Linked Heading',
		basedOn: 'Heading1',
		paragraphFormatting: { spacingBeforeTwips: twips(160) },
		runFormatting: { bold: true },
	},
	{
		kind: 'table',
		name: 'Notes Table',
		paragraphFormatting: { align: 'right' },
		runFormatting: { bold: true },
		tableProperties: {
			shadingFill: '#FFCC99',
			cellMargins: { left: twips(200) },
			conditional: { firstRow: { run: { italic: true } } },
		},
	},
] satisfies StyleEdit[];

describe('Word style catalog saves', () => {
	it('authors all four style kinds and saves their properties', async () => {
		let model = createDocument();
		for (const edit of createStyles) {
			model = { ...model, ...must(editStyleCatalogs(model, edit)).catalogs };
		}
		const { model: reopened } = await loadDocx(await saveDocx(model));
		expect(reopened.paragraphStyles?.styles.ArticleHeading).toMatchObject({
			name: 'Article Heading',
			basedOn: 'Normal',
			next: 'Normal',
			uiPriority: 7,
			quickFormat: false,
			customStyle: true,
			formatting: { align: 'center' },
		});
		expect(reopened.characterStyles?.styles.ArticleHeading?.formatting).toEqual({ fontSize: 16 });
		expect(reopened.characterStyles?.styles.Gloss).toMatchObject({
			type: 'character',
			formatting: { italic: true, fontFamily: 'Georgia' },
		});
		expect(reopened.paragraphStyles?.styles.LinkedHeading).toMatchObject({
			linkedStyle: 'LinkedHeadingChar',
			basedOn: 'Heading1',
		});
		expect(reopened.characterStyles?.styles.LinkedHeadingChar).toMatchObject({
			type: 'character',
			linkedStyle: 'LinkedHeading',
			formatting: { bold: true },
		});
		expect(reopened.tableStyles?.styles.NotesTable).toMatchObject({
			paragraphFormatting: { align: 'right' },
			runFormatting: { bold: true },
			shadingFill: '#FFCC99',
			cellMargins: { left: 200 },
			conditional: { firstRow: { run: { italic: true } } },
		});
	});

	it('patches both source catalogs and latent exceptions while keeping extension XML', async () => {
		const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
		const xml = must(await zip.file('word/styles.xml')?.async('string'))
			.replace('<w:styles ', '<w:styles xmlns:x="urn:test" ')
			.replace(
				'<w:style ',
				'<w:latentStyles w:count="2" w:defUIPriority="9" w:defQFormat="0" x:keep="default">' +
					'<w:lsdException w:name="Heading 1" w:semiHidden="1" x:keep="exception"/><x:unknown/></w:latentStyles><w:style ',
			)
			.replace('<w:name w:val="Normal"/>', '<w:name w:val="Normal"/><x:keep/><x:extLst/>')
			.replace('</w:styles>', '<x:extLst/></w:styles>');
		zip.file('word/styles.xml', xml);
		zip.file('word/stylesWithEffects.xml', xml.replace('<x:keep/>', '<x:keep x:effects="1"/>'));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = loaded.model;
		const latent = must(model.latentStyles);
		expect(latent).toEqual({
			count: 2,
			defaults: { uiPriority: 9, qFormat: false },
			exceptions: [{ name: 'Heading 1', semiHidden: true }],
		});
		const paragraph = must(model.paragraphStyles?.styles.Normal);
		const edited = must(
			editStyleCatalogs(model, {
				kind: 'paragraph',
				id: 'Normal',
				name: 'Page Body',
				next: 'Heading1',
				paragraphFormatting: { ...paragraph.formatting, align: 'right' },
				runFormatting: { italic: true },
			}),
		);
		const renamedCharacter = must(
			editStyleCatalogs(edited.catalogs, {
				kind: 'character',
				id: 'Hyperlink',
				name: 'Web Link',
				basedOn: 'DefaultParagraphFont',
				runFormatting: { ...model.characterStyles?.styles.Hyperlink?.formatting, bold: true },
			}),
		);
		const editedTable = must(
			editStyleCatalogs(renamedCharacter.catalogs, {
				kind: 'table',
				id: 'TableNormal',
				name: 'Document Table',
				paragraphFormatting: {},
				runFormatting: { italic: true },
				tableProperties: { shadingFill: '#123456' },
			}),
		);
		const next = {
			...model,
			...editedTable.catalogs,
			latentStyles: {
				...latent,
				defaults: { uiPriority: 10, locked: true },
				exceptions: [
					{ name: 'Heading 1', semiHidden: false },
					{ name: 'Title', qFormat: true },
				],
			},
		};
		const saved = await loaded.save(next);
		const reopened = await loadDocx(saved);
		expect(reopened.model.latentStyles).toEqual(next.latentStyles);
		expect(reopened.model.paragraphStyles?.styles.Normal).toMatchObject({
			name: 'Page Body',
			next: 'Heading1',
			formatting: { align: 'right' },
		});
		expect(reopened.model.characterStyles?.styles.Normal?.formatting.italic).toBe(true);
		expect(reopened.model.characterStyles?.styles.Hyperlink).toMatchObject({
			name: 'Web Link',
			formatting: { bold: true, underline: true },
		});
		expect(reopened.model.tableStyles?.styles.TableNormal).toMatchObject({
			name: 'Document Table',
			runFormatting: { italic: true },
			shadingFill: '#123456',
			cellMargins: { left: 108 },
		});
		const output = await JSZip.loadAsync(saved);
		for (const path of ['word/styles.xml', 'word/stylesWithEffects.xml']) {
			const part = must(await output.file(path)?.async('string'));
			expect(part).toContain('Page Body');
			expect(part).toContain(
				path.endsWith('Effects.xml') ? '<x:keep x:effects="1"/>' : '<x:keep/>',
			);
			expect(part.match(/<x:extLst\/>/g)).toHaveLength(2);
			expect(part).toContain('x:keep="exception"');
			expect(part).toContain('<x:unknown/>');
		}
	});

	it('keeps independently authored linked partner formatting and rejects inheritance cycles', () => {
		const model = createDocument();
		const created = must(editStyleCatalogs(model, must(createStyles[2])));
		const runStyles = must(created.catalogs.characterStyles).styles;
		const partner = must(runStyles.LinkedHeadingChar);
		runStyles[partner.id] = {
			...partner,
			name: 'Custom partner name',
			formatting: { fontSize: 24, bold: true },
		};
		const changed = must(
			editStyleCatalogs(created.catalogs, {
				kind: 'linked',
				id: 'LinkedHeading',
				name: 'New heading',
				basedOn: 'Heading1',
				paragraphFormatting: {},
				runFormatting: { bold: true, italic: true },
			}),
		);
		expect(changed.catalogs.characterStyles?.styles.LinkedHeadingChar).toMatchObject({
			name: 'Custom partner name',
			formatting: { fontSize: 24, bold: true, italic: true },
		});
		expect(
			editStyleCatalogs(changed.catalogs, {
				kind: 'paragraph',
				id: 'Heading1',
				name: 'Heading 1',
				basedOn: 'LinkedHeading',
				paragraphFormatting: {},
				runFormatting: {},
			}),
		).toBeUndefined();
	});

	it('allocates style ids independently of source-only numbering styles', async () => {
		const model = createDocument();
		const edit = must(
			editStyleCatalogs(model, { kind: 'character', name: 'No List', runFormatting: {} }),
		);
		expect(edit.id).toBe('NoList2');
		const zip = await JSZip.loadAsync(await saveDocx({ ...model, ...edit.catalogs }));
		const xml = await zip.file('word/styles.xml')?.async('string');
		expect(xml).toContain('w:type="numbering" w:default="1" w:styleId="NoList"');
		expect(xml).toContain('w:styleId="NoList2"');
	});
});
