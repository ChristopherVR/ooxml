import JSZip from 'jszip';
import { expect, test } from 'vitest';
import { createDocument, loadDocx, saveDocx, type Paragraph } from '../index';
import { children, first, getW, parseXml } from '../xml';
import { bibliographyGallery, buildReferenceTable, type Bibliography } from './index';

test.each(['Bibliography', 'Works Cited'])(
	'writes the %s gallery as a styled block control',
	async (heading) => {
		const model = createDocument();
		const bibliography: Bibliography = {
			sources: [{ tag: 'Doe20', type: 'Book', fields: { Title: 'Field studies' } }],
		};
		const result = buildReferenceTable(model, 'BIBLIOGRAPHY', { bibliography, empty: 'None' });
		bibliographyGallery(model, '123', heading, result);
		const blocks: Paragraph[] = result.map((p, i) => ({ ...p, id: 'gallery-' + i }));
		const bytes = await saveDocx({ ...model, blocks });
		const zip = await JSZip.loadAsync(bytes);
		const xml = parseXml(await zip.file('word/document.xml')!.async('text'));
		const body = first(xml.documentElement, 'body')!;
		const [control] = children(body, 'sdt');
		const gallery = first(first(first(control!, 'sdtPr'), 'docPartObj'), 'docPartGallery');
		expect(getW(gallery, 'val')).toBe('Bibliographies');
		const paragraphs = children(first(control!, 'sdtContent')!, 'p');
		expect(paragraphs.map((p) => getW(first(first(p, 'pPr'), 'pStyle'), 'val'))).toEqual([
			'Heading1',
			'Bibliography',
			undefined,
		]);
		const reopened = await loadDocx(bytes);
		expect(reopened.model.blocks.map((block) => block.contentControl?.id)).toEqual([
			'123',
			'123',
			'123',
		]);
		expect(await reopened.save()).toEqual(bytes);
	},
);
