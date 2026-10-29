import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { expectParagraph, must } from './test-support/access';
import { WORD_NS } from './xml';

const STORE = '{11111111-2222-3333-4444-555555555555}';
const ITEM = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ns0:root xmlns:ns0="urn:x" id="a"><ns0:name>Ada</ns0:name><ns0:tags><ns0:tag>one</ns0:tag><ns0:tag>two</ns0:tag></ns0:tags></ns0:root>`;
const PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ds:datastoreItem ds:itemID="${STORE}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`;

const control = (id: number, xpath: string, text: string, mappings = "xmlns:ns0='urn:x'") =>
	`<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:dataBinding w:prefixMappings="${mappings}" w:xpath="${xpath}" w:storeItemID="${STORE}"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;

async function load(body: string, withStore = true) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p>${body}</w:p><w:sectPr/></w:body></w:document>`,
	);
	if (withStore) {
		zip.file('customXml/item1.xml', ITEM);
		zip.file('customXml/itemProps1.xml', PROPS);
	}
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, paragraph: expectParagraph(loaded.model.blocks[0]) };
}
const itemOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('customXml/item1.xml')?.async('string')) ?? '';

describe('data-bound content controls', () => {
	it('writes an edited value to the element the XPath selects', async () => {
		const { loaded, paragraph } = await load(control(1, '/ns0:root/ns0:name', 'Ada'));
		must(paragraph.runs[0]).text = 'Grace';
		const item = await itemOf(await loaded.save(loaded.model));
		expect(item).toContain('<ns0:name>Grace</ns0:name>');
		expect(item).toContain('<ns0:tag>two</ns0:tag>');
	});

	it('follows an index step and an attribute step', async () => {
		const { loaded, paragraph } = await load(
			control(1, '/ns0:root/ns0:tags/ns0:tag[2]', 'two') + control(2, '/ns0:root/@id', 'a'),
		);
		for (const run of paragraph.runs) run.text = run.text === 'two' ? '2nd' : 'b';
		const item = await itemOf(await loaded.save(loaded.model));
		expect(item).toContain('<ns0:tag>one</ns0:tag><ns0:tag>2nd</ns0:tag>');
		expect(item).toContain('id="b"');
	});

	it('leaves the part alone when the bound text did not change', async () => {
		const { loaded, paragraph } = await load(
			control(1, '/ns0:root/ns0:name', 'Ada') + '<w:r><w:t>tail</w:t></w:r>',
		);
		must(paragraph.runs.at(-1)).text = ' changed';
		expect(await itemOf(await loaded.save(loaded.model))).toBe(ITEM);
	});

	it('refuses an XPath it cannot evaluate, or one that selects nothing', async () => {
		const unsupported = await load(control(1, '//ns0:name[@x=&quot;1&quot;]', 'Ada'));
		must(unsupported.paragraph.runs[0]).text = 'Grace';
		await expect(unsupported.loaded.save(unsupported.loaded.model)).rejects.toThrow(
			/data-bound content control 1.*not supported/,
		);
		const missing = await load(control(1, '/ns0:root/ns0:absent', 'Ada'));
		must(missing.paragraph.runs[0]).text = 'Grace';
		await expect(missing.loaded.save(missing.loaded.model)).rejects.toThrow(/selects no node/);
	});

	it('saves a bound control whose store part is not in the package', async () => {
		const { loaded, paragraph } = await load(control(1, '/ns0:root/ns0:name', 'Ada'), false);
		must(paragraph.runs[0]).text = 'Grace';
		const bytes = await loaded.save(loaded.model);
		expect(
			await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string'),
		).toContain('Grace');
	});
});
