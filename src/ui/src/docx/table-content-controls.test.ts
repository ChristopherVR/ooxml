import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DOMParser, DOMSerializer } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
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

	it('takes a control from pasted HTML only when it has the shape of one', () => {
		// A page can put anything in data-sdt: what is not a control (an object with an id and a
		// kind) is not taken, and one that is takes the level of where it is.
		const html = document.createElement('div');
		html.innerHTML =
			`<p><span data-sdt='5'>a</span><span data-sdt='{"id":"7"}'>b</span>` +
			`<span data-sdt='{"id":"8","kind":"text","level":"cell"}'>c</span></p>` +
			`<table><tbody><tr><td data-sdt='[1]'><p>d</p></td><td data-sdt='{"kind":"text"}'><p>e</p></td>` +
			`<td data-sdt='{"id":"9","kind":"checkbox","level":"inline"}'><p>f</p></td></tr></tbody></table>`;
		const doc = DOMParser.fromSchema(schema).parse(html);
		expect(
			controlSpans(doc).map((span) => [span.level, span.control.id, span.control.level]),
		).toEqual([
			['inline', '8', 'inline'],
			['cell', '9', 'cell'],
		]);
	});

	it('keeps a cell control through HTML, as when cells are copied and pasted', async () => {
		const doc = modelToDoc((await form()).model);
		const html = document.createElement('div');
		html.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
		const parsed = DOMParser.fromSchema(schema).parse(html);
		expect(controlSpans(parsed)).toEqual(controlSpans(doc));
		expect(controlSpans(parsed)[0]).toMatchObject({ level: 'cell', control: { id: '71' } });
	});

	it('keeps the verbatim properties of a pasted cell control only when they come from the editor', async () => {
		const bound = (path: string) => `<w:dataBinding w:xpath="${path}" w:storeItemID="{1}"/>`;
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/>` +
				`</w:tblGrid><w:tr><w:sdt><w:sdtPr><w:id w:val="71"/>${bound('/mine')}</w:sdtPr>` +
				`<w:sdtContent>${cell('Mine')}</w:sdtContent></w:sdt></w:tr></w:tbl>` +
				`<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const doc = modelToDoc(loaded.model);
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({ schema, doc, plugins: [contentControlPlugin()] }),
		});
		const paste = (html: string) => {
			let after = 0;
			view.state.doc.forEach((node, offset) => {
				if (node.textContent === 'After') after = offset + node.nodeSize - 1;
			});
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, after)));
			// The paste reads nothing else from the event.
			view.pasteHTML(html, new Event('paste') as ClipboardEvent);
		};
		// A table from another page: properties that bind its cell to that page's data go.
		const td = document.createElement('td');
		td.dataset.sdt = JSON.stringify({
			id: '90',
			level: 'cell',
			kind: 'richText',
			alias: 'Theirs',
			sourceXml: `<w:sdtPr xmlns:w="${W}"><w:alias w:val="Theirs"/><w:id w:val="90"/>${bound('/theirs')}</w:sdtPr>`,
			endSourceXml: `<w:sdtEndPr xmlns:w="${W}"><w:rPr><w:b/></w:rPr></w:sdtEndPr>`,
		});
		td.innerHTML = '<p>Theirs</p>';
		paste(`<table data-pm-slice="0 0 []"><tbody><tr>${td.outerHTML}</tr></tbody></table>`);
		const controlOf = (id: string) =>
			controlSpans(view.state.doc).find((span) => span.level === 'cell' && span.control.id === id);
		expect(controlOf('90')?.control).toEqual({
			id: '90',
			level: 'cell',
			kind: 'richText',
			alias: 'Theirs',
		});
		const bytes = await loaded.save(docToModel(view.state.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:t>Theirs</w:t>');
		expect(xml).toContain('<w:alias w:val="Theirs"/><w:id w:val="90"/>');
		expect(xml).toContain('w:xpath="/mine"');
		expect(xml).not.toContain('/theirs');
		expect(xml).not.toContain('w:sdtEndPr');
		// A copy of the document's own table keeps its cell's properties, under a new id.
		const own = controlOf('71')!.control.sourceXml;
		paste(
			view.serializeForClipboard(view.state.doc.slice(0, doc.firstChild!.nodeSize)).dom.innerHTML,
		);
		const copies = controlSpans(view.state.doc).filter(
			(span) => span.level === 'cell' && span.control.sourceXml === own,
		);
		expect(own).toContain('w:xpath="/mine"');
		expect(copies.map((span) => span.control.id === '71')).toEqual([true, false]);
		view.destroy();
		host.remove();
	});

	it('keeps no properties another page gives a table or a row through the slice context', async () => {
		// Pasted HTML names the nodes around what it holds, and their attributes, in data-pm-slice:
		// a page can put a control there on a table or a row, which no parse rule reads.
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const doc = modelToDoc(loaded.model);
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({ schema, doc, plugins: [contentControlPlugin()] }),
		});
		const theirs = (id: string, level: string) => ({
			id,
			level,
			kind: 'richText',
			sourceXml: `<w:sdtPr xmlns:w="${W}"><w:id w:val="${id}"/><w:dataBinding w:xpath="/theirs" w:storeItemID="{2}"/></w:sdtPr>`,
		});
		const context = (nodes: unknown[]) => JSON.stringify(nodes).replace(/'/g, '&#39;');
		for (const html of [
			`<tr data-pm-slice='0 0 ${context(['table', { contentControl: theirs('56', 'block') }])}'><td><p>Table</p></td></tr>`,
			`<td data-pm-slice='0 0 ${context(['table', null, 'tableRow', { properties: { contentControl: theirs('57', 'row') } }])}'><p>Row</p></td>`,
		]) {
			let after = 0;
			view.state.doc.forEach((node, offset) => {
				if (node.textContent === 'After') after = offset + node.nodeSize - 1;
			});
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, after)));
			// The paste reads nothing else from the event.
			view.pasteHTML(html, new Event('paste') as ClipboardEvent);
		}
		const held: unknown[] = [];
		view.state.doc.descendants((node) => {
			const control = node.attrs.contentControl ?? node.attrs.properties?.contentControl;
			if (control && node.type.name !== 'tableCell') held.push(control);
		});
		expect(held).toHaveLength(2);
		expect(held).toEqual(
			expect.arrayContaining([
				{ id: '56', level: 'block', kind: 'richText' },
				{ id: '57', level: 'row', kind: 'richText' },
			]),
		);
		const bytes = await loaded.save(docToModel(view.state.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:sdt><w:sdtPr><w:id w:val="57"/></w:sdtPr><w:sdtContent><w:tr>');
		expect(xml).not.toContain('/theirs');
		view.destroy();
		host.remove();
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

	it('applies the rules of a cell control to a paste, as to typing', async () => {
		// Typing over a cell's placeholder replaces it and the control no longer shows one, and a list
		// or check box cell takes no typing: a paste is input too, as plain text or as HTML.
		const prompt = (id: string, type: string, text: string) =>
			`<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/>${type}</w:sdtPr><w:sdtContent>` +
			`<w:tc><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t>` +
			`</w:r></w:p></w:tc></w:sdtContent></w:sdt>`;
		const list = '<w:dropDownList><w:listItem w:displayText="A" w:value="a"/></w:dropDownList>';
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/>` +
				`<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>` +
				prompt('81', '<w:text/>', 'Click or tap here to enter text.') +
				prompt('82', '', 'Click or tap here to enter text.') +
				prompt('83', list, 'Choose an item.') +
				prompt(
					'84',
					'<w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>',
					'\u2610',
				) +
				`</w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const doc = modelToDoc(loaded.model);
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({ schema, doc, plugins: [contentControlPlugin()] }),
		});
		const spanOf = (id: string) =>
			controlSpans(view.state.doc).find((span) => span.level === 'cell' && span.control.id === id)!;
		// The caret after the first character of the cell's text.
		const into = (id: string) =>
			view.dispatch(
				view.state.tr.setSelection(TextSelection.create(view.state.doc, spanOf(id).from + 3)),
			);
		// The paste reads nothing else from the event.
		const event = () => new Event('paste') as ClipboardEvent;
		into('81');
		view.pasteText('DOC-7', event());
		into('82');
		view.pasteHTML('<p><b>Second</b> sheet</p>', event());
		const filled = (id: string) => view.state.doc.nodeAt(spanOf(id).from)!;
		expect(filled('81').textContent).toBe('DOC-7');
		expect(filled('82').textContent).toBe('Second sheet');
		for (const id of ['81', '82'])
			expect(spanOf(id).control.showingPlaceholder, id).toBeUndefined();
		const before = view.state.doc;
		for (const id of ['83', '84']) {
			into(id);
			view.pasteText('Other', event());
			view.pasteHTML('<p>Other</p>', event());
		}
		expect(view.state.doc.eq(before)).toBe(true);
		const bytes = await loaded.save(docToModel(view.state.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml.match(/<w:showingPlcHdr\/>/g)).toHaveLength(2);
		expect(xml).toContain('<w:id w:val="81"/><w:text/></w:sdtPr>');
		expect(xml).toContain('<w:t>DOC-7</w:t>');
		expect(xml).not.toMatch(/PlaceholderText"\/><\/w:rPr><w:t>DOC-7/);
		view.destroy();
		host.remove();
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

	for (const [around, edited] of [
		['a paragraph, a table and a paragraph', 'Inside'],
		['a paragraph, a table and a paragraph', 'After'],
		['a table', 'Inside'],
	] as const)
		it(`removes a temporary control around ${around} once ${edited} is edited`, async () => {
			const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
			const table =
				`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
				`<w:tr>${cell('Inside')}</w:tr></w:tbl>`;
			const content =
				around === 'a table' ? table : paragraph('Before') + table + paragraph('After');
			const zip = new JSZip();
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtPr><w:alias w:val="Section"/>` +
					`<w:id w:val="76"/><w:temporary/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>` +
					`${paragraph('Elsewhere')}<w:sectPr/></w:body></w:document>`,
			);
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const doc = modelToDoc(loaded.model);
			const state = EditorState.create({ schema, doc, plugins: [contentControlPlugin()] });
			const end = (s: EditorState, text: string) => {
				let found = -1;
				s.doc.descendants((node, pos) => {
					if (found < 0 && node.isText && node.text === text) found = pos + node.nodeSize;
				});
				return found;
			};
			const save = async (s: EditorState) => {
				const bytes = await loaded.save(docToModel(s.doc, loaded.model, doc));
				return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			};
			// An edit outside it leaves it; an edit in it removes all of it.
			const outside = state.apply(state.tr.insertText('!', end(state, 'Elsewhere')));
			expect(await save(outside)).toContain('<w:alias w:val="Section"/>');
			const next = outside.apply(outside.tr.insertText('!', end(outside, edited)));
			const xml = await save(next);
			expect(xml).toContain(`<w:t>${edited}!</w:t>`);
			expect(xml).not.toContain('<w:sdt>');
		});

	for (const temporary of [false, true])
		it(`${temporary ? 'drops a temporary' : 'keeps a'} block control around a table the editor holds none for`, async () => {
			const zip = new JSZip();
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${W}"><w:body><w:sdt><w:sdtPr><w:alias w:val="Sheet"/><w:id w:val="77"/>` +
					`${temporary ? '<w:temporary/>' : ''}</w:sdtPr><w:sdtContent><w:tbl><w:tblPr/>` +
					`<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell('Title')}</w:tr></w:tbl>` +
					`</w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>`,
			);
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const doc = modelToDoc(loaded.model);
			expect(doc.firstChild!.attrs.contentControl).toMatchObject({ id: '77' });
			// As a document shared with an editor that kept no controls on tables has it: only a
			// temporary control is ever taken off a table, so another one stays.
			const state = EditorState.create({ schema, doc });
			const lost = state.tr.setNodeMarkup(0, undefined, {
				...doc.firstChild!.attrs,
				contentControl: null,
			});
			const bytes = await loaded.save(docToModel(state.apply(lost).doc, loaded.model, doc));
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			if (temporary) expect(xml).not.toContain('<w:sdt>');
			else
				expect(xml).toMatch(/<w:body><w:sdt><w:sdtPr><w:alias w:val="Sheet"\/><w:id w:val="77"\/>/);
		});

	for (const merged of [false, true])
		for (const temporary of [false, true])
			it(`${temporary ? 'drops a temporary' : 'keeps a'} cell control the editor holds none for${merged ? ', in a merged table' : ''}`, async () => {
				const loaded = await form(temporary ? '<w:temporary/>' : '', merged);
				const doc = modelToDoc(loaded.model);
				// As a document shared with an editor that kept no controls on cells has it: only a
				// temporary control is ever taken off a cell, so another one stays.
				const state = EditorState.create({ schema, doc });
				let at = -1;
				state.doc.descendants((node, pos) => {
					if (at < 0 && node.type.name === 'tableCell' && node.attrs.contentControl) at = pos;
				});
				const lost = state.tr.setNodeMarkup(at, undefined, {
					...state.doc.nodeAt(at)!.attrs,
					contentControl: null,
				});
				const bytes = await loaded.save(docToModel(state.apply(lost).doc, loaded.model, doc));
				const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
				if (temporary) expect(xml).not.toContain('<w:sdt>');
				else expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="Kind"/><w:id w:val="71"/>');
			});

	it('gives a cell added before a cell control none of its control', async () => {
		// A cell keeps the control of the cell it was loaded as, found by its source cell, not by
		// its place: a cell added before it, as Insert Column adds one, has none.
		const loaded = await form();
		const doc = modelToDoc(loaded.model);
		const state = EditorState.create({ schema, doc });
		let at = -1;
		state.doc.descendants((node, pos) => {
			if (at < 0 && node.type.name === 'tableCell' && node.attrs.contentControl) at = pos;
		});
		const added = schema.nodes.tableCell.create(null, schema.nodes.paragraph.create({ id: 'n1' }));
		const next = state.apply(state.tr.insert(at, added));
		const bytes = await loaded.save(docToModel(next.doc, loaded.model, doc));
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
		expect(xml).toContain(
			'<w:sdtContent><w:tc><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/>',
		);
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
