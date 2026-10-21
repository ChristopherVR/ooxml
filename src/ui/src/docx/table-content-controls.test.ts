import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DOMParser, DOMSerializer } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { loadDocx } from 'ooxml-core/docx';
import { chooseListItem, contentControlPlugin, controlSpans } from 'ooxml-core/docx/ui';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (text: string, properties = '') =>
	`<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ''}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const choice = (extra: string) =>
	`<w:sdt><w:sdtPr><w:alias w:val="Kind"/><w:id w:val="71"/>${extra}<w:showingPlcHdr/>` +
	`<w:dropDownList><w:listItem w:displayText="Primary" w:value="primary"/>` +
	`<w:listItem w:displayText="Secondary" w:value="secondary"/></w:dropDownList></w:sdtPr>` +
	`<w:sdtContent><w:tc><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>` +
	`<w:t>Choose an item.</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt>`;
/** A table of sources: a title and a kind picked from a list, the title merged down a row. */
async function form(extra = '', merged = false) {
	const zip = new JSZip();
	const title = merged ? cell('First source', '<w:vMerge w:val="restart"/>') : cell('First source');
	const below = merged ? `<w:tr>${cell('', '<w:vMerge/>')}${cell('Second sheet')}</w:tr>` : '';
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/>` +
			`<w:gridCol w:w="3000"/></w:tblGrid><w:tr>${title}${choice(extra)}</w:tr>${below}` +
			`</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
async function pick(loaded: Awaited<ReturnType<typeof form>>, index: number) {
	const doc = modelToDoc(loaded.model);
	const state = EditorState.create({ schema, doc });
	const next = state.apply(chooseListItem(state, '71', index)!);
	const model = docToModel(next.doc, loaded.model, doc);
	const bytes = await loaded.save(model);
	return {
		next,
		xml: await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string'),
	};
}

describe('content controls around table cells in the editor', () => {
	it('draws a cell control with its kind, title and placeholder', async () => {
		const loaded = await form();
		const doc = modelToDoc(loaded.model);
		const [span] = controlSpans(doc);
		expect(span).toMatchObject({ level: 'cell', control: { id: '71', alias: 'Kind' } });
		const td = DOMSerializer.fromSchema(schema).serializeNode(
			doc.nodeAt(span!.from)!,
		) as HTMLElement;
		expect(td.dataset).toMatchObject({
			sdtCell: '71',
			sdtKind: 'dropDownList',
			sdtTitle: 'Kind',
			sdtPlaceholder: '1',
		});
		expect(JSON.parse(td.dataset.sdt!).items).toHaveLength(2);
	});

	it('writes a cell control pasted from another page without content in its properties', async () => {
		const loaded = await form();
		const doc = modelToDoc(loaded.model);
		// The page put a field that would fetch a picture where the control's properties go.
		const control = {
			id: '90',
			level: 'cell',
			kind: 'richText',
			sourceXml: `<w:p xmlns:w="${W}"><w:fldSimple w:instr='INCLUDEPICTURE "http://example.test/a.png"'/></w:p>`,
		};
		const td = document.createElement('td');
		td.dataset.sdt = JSON.stringify(control);
		td.append(Object.assign(document.createElement('p'), { textContent: 'Pasted' }));
		const [html, table, body, row] = ['div', 'table', 'tbody', 'tr'].map((tag) =>
			document.createElement(tag),
		);
		row!.append(td);
		body!.append(row!);
		table!.append(body!);
		html!.append(table!);
		const pasted = DOMParser.fromSchema(schema).parse(html!);
		const state = EditorState.create({ schema, doc });
		const next = state.apply(state.tr.insert(doc.content.size, pasted.content));
		expect(controlSpans(next.doc).map((span) => span.control.id)).toContain('90');
		const bytes = await loaded.save(docToModel(next.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:t>Pasted</w:t>');
		expect(xml).toMatch(/<w:sdt><w:sdtPr><w:id w:val="90"\/><\/w:sdtPr><w:sdtContent><w:tc>/);
		expect(xml).not.toContain('INCLUDEPICTURE');
	});

	it('keeps a cell control through HTML, as when cells are copied and pasted', async () => {
		const doc = modelToDoc((await form()).model);
		const html = document.createElement('div');
		html.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
		const parsed = DOMParser.fromSchema(schema).parse(html);
		expect(controlSpans(parsed)).toEqual(controlSpans(doc));
		expect(controlSpans(parsed)[0]).toMatchObject({ level: 'cell', control: { id: '71' } });
	});

	for (const merged of [false, true])
		it(`saves an entry picked in a cell ${merged ? 'of a merged table ' : ''}without its placeholder flag`, async () => {
			const { next, xml } = await pick(await form('', merged), 1);
			expect(next.doc.textContent).toContain('Secondary');
			expect(xml).toContain('<w:alias w:val="Kind"/><w:id w:val="71"/><w:dropDownList>');
			expect(xml).not.toContain('showingPlcHdr');
			expect(xml).toContain('<w:t>Secondary</w:t>');
			expect(xml).not.toContain('PlaceholderText');
		});

	it('saves an entry picked in a temporary cell control as the cell text alone', async () => {
		const { xml } = await pick(await form('<w:temporary/>'), 0);
		expect(xml).not.toContain('<w:sdt');
		expect(xml).toContain('<w:tc><w:p><w:r><w:t>Primary</w:t></w:r></w:p></w:tc>');
	});

	for (const merged of [false, true])
		it(`keeps a row control when text is typed in its ${merged ? 'merged ' : ''}table`, async () => {
			const zip = new JSZip();
			const first = merged ? cell('Title', '<w:vMerge w:val="restart"/>') : cell('Title');
			const continued = merged ? cell('', '<w:vMerge/>') : cell('Again');
			const item =
				`<w:sdt><w:sdtPr><w:alias w:val="Item"/><w:id w:val="72"/></w:sdtPr><w:sdtContent>` +
				`<w:tr>${continued}${cell('Line')}</w:tr></w:sdtContent></w:sdt>`;
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/>` +
					`<w:gridCol w:w="3000"/></w:tblGrid><w:tr>${first}${cell('Value')}</w:tr>${item}` +
					`</w:tbl><w:sectPr/></w:body></w:document>`,
			);
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const doc = modelToDoc(loaded.model);
			const state = EditorState.create({ schema, doc });
			let at = -1;
			state.doc.descendants((node, pos) => {
				if (at < 0 && node.isText && node.text === 'Line') at = pos + node.nodeSize;
			});
			const next = state.apply(state.tr.insertText(' two', at));
			const bytes = await loaded.save(docToModel(next.doc, loaded.model, doc));
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			expect(xml).toContain('<w:t>Line two</w:t>');
			expect(xml).toMatch(
				/<w:sdt><w:sdtPr><w:alias w:val="Item"\/><w:id w:val="72"\/><\/w:sdtPr><w:sdtContent><w:tr>/,
			);
		});

	for (const merged of [false, true])
		it(`removes a temporary row control once its ${merged ? 'merged ' : ''}row is edited`, async () => {
			const zip = new JSZip();
			const first = merged ? cell('Title', '<w:vMerge w:val="restart"/>') : cell('Title');
			const continued = merged ? cell('', '<w:vMerge/>') : cell('Again');
			const item =
				`<w:sdt><w:sdtPr><w:alias w:val="Item"/><w:id w:val="73"/><w:temporary/></w:sdtPr>` +
				`<w:sdtContent><w:tr>${continued}${cell('Line')}</w:tr></w:sdtContent></w:sdt>`;
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/>` +
					`<w:gridCol w:w="3000"/></w:tblGrid><w:tr>${first}${cell('Value')}</w:tr>${item}` +
					`</w:tbl><w:sectPr/></w:body></w:document>`,
			);
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const doc = modelToDoc(loaded.model);
			const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
			const at = (s: EditorState, text: string) => {
				let found = -1;
				s.doc.descendants((node, pos) => {
					if (found < 0 && node.isText && node.text === text) found = pos + node.nodeSize;
				});
				return found;
			};
			// An edit in another row leaves the control; an edit in its row removes it.
			const elsewhere = state.apply(state.tr.insertText('!', at(state, 'Value')));
			expect(elsewhere.doc.firstChild!.child(1).attrs.properties?.contentControl?.id).toBe('73');
			const next = elsewhere.apply(elsewhere.tr.insertText(' two', at(elsewhere, 'Line')));
			expect(next.doc.firstChild!.child(1).attrs.properties?.contentControl).toBeUndefined();
			const bytes = await loaded.save(docToModel(next.doc, loaded.model, doc));
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			expect(xml).toContain('<w:t>Line two</w:t>');
			expect(xml).not.toContain('<w:sdt>');
		});

	it('keeps a block control around a table, and one around a cell paragraph, through an edit', async () => {
		const zip = new JSZip();
		const note =
			`<w:tc><w:sdt><w:sdtPr><w:alias w:val="Note"/><w:id w:val="75"/></w:sdtPr><w:sdtContent>` +
			`<w:p><w:r><w:t>Remark</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc>`;
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtPr><w:alias w:val="Sheet"/><w:id w:val="74"/>` +
				`</w:sdtPr><w:sdtContent><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/>` +
				`<w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('Title')}${note}</w:tr></w:tbl>` +
				`</w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const doc = modelToDoc(loaded.model);
		const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
		let next = state;
		for (const word of ['Title', 'Remark']) {
			let at = -1;
			next.doc.descendants((node, pos) => {
				if (at < 0 && node.isText && node.text === word) at = pos + node.nodeSize;
			});
			next = next.apply(next.tr.insertText(' two', at));
		}
		const bytes = await loaded.save(docToModel(next.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:t>Title two</w:t>');
		expect(xml).toContain('<w:t>Remark two</w:t>');
		expect(xml).toMatch(/<w:body><w:sdt><w:sdtPr><w:alias w:val="Sheet"\/><w:id w:val="74"\/>/);
		expect(xml).toMatch(/<w:tc><w:sdt><w:sdtPr><w:alias w:val="Note"\/><w:id w:val="75"\/>/);
	});
});
