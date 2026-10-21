// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx } from 'ooxml-core/docx';
import {
	contentControlPlugin,
	controlSpans,
	typeOverPlaceholder,
	updateControlProperties,
} from 'ooxml-core/docx/ui';
import { contentControlViewPlugin } from './content-control-view';
import { parseListEntries } from './content-control-dialog';
import { docToModel, modelToDoc } from './model-adapter';

// Vitest runs from src/ui; the Word-authored fixture lives with the core tests.
const fixture = resolve(
	process.cwd(),
	'../core/docx/__fixtures__/content-controls/content-controls.docx',
);
async function open() {
	const loaded = await loadDocx(new Uint8Array(await readFile(fixture)));
	const doc = modelToDoc(loaded.model);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			plugins: [contentControlPlugin(), contentControlViewPlugin()],
		}),
	});
	return { loaded, view, host };
}
const savedXml = async (loaded: Awaited<ReturnType<typeof loadDocx>>, view: EditorView) => {
	const bytes = await loaded.save(docToModel(view.state.doc, loaded.model));
	return (await (
		await JSZip.loadAsync(bytes)
	)
		.file('word/document.xml')!
		.async('string')) as string;
};
const click = (view: EditorView, element: Element) =>
	view.someProp('handleClick', (handler) =>
		handler(view, 0, { target: element, clientX: 0 } as unknown as MouseEvent),
	);

describe('content controls in the editor', () => {
	it('draws every control with its kind, title and placeholder state', async () => {
		const { view, host } = await open();
		const spans = [...host.querySelectorAll<HTMLElement>('.dve-sdt')];
		expect(spans.map((span) => span.dataset.sdtKind)).toEqual([
			'text',
			'checkbox',
			'dropDownList',
			'comboBox',
			'date',
		]);
		expect(spans[0]?.dataset.sdtTitle).toBe('Full name');
		expect(spans[0]?.dataset.sdtPlaceholder).toBe('1');
		expect(spans[0]?.dataset.sdtLock).toBe('sdtLocked');
		expect(spans[2]?.dataset.sdtPlaceholder).toBeUndefined();
		const block = host.querySelector<HTMLElement>('p[data-sdt-block]');
		expect(block?.dataset.sdtTitle).toBe('Notes block');
		expect(block?.dataset.sdtLock).toBe('contentLocked');
		view.destroy();
	});

	it('toggles a check box on click and saves the new state and glyph', async () => {
		const { loaded, view, host } = await open();
		const box = host.querySelector<HTMLElement>('.dve-sdt[data-sdt-kind="checkbox"]')!;
		expect(box.textContent).toBe('☒');
		expect(click(view, box)).toBe(true);
		expect(view.state.doc.textContent).toContain('Agreed: ☐');
		const xml = await savedXml(loaded, view);
		expect(xml).toContain('<w14:checked w14:val="0"/>');
		expect(xml).toContain('<w:t>☐</w:t>');
		view.destroy();
	});

	it('picks a drop-down entry and a date, saving both', async () => {
		const { loaded, view, host } = await open();
		const drop = host.querySelector<HTMLElement>('.dve-sdt[data-sdt-kind="dropDownList"]')!;
		expect(click(view, drop)).toBe(true);
		const options = [...document.querySelectorAll<HTMLElement>('.ribbon-list [role="option"]')];
		expect(options.map((option) => option.textContent)).toEqual(['Red', 'Green', 'Blue']);
		options[2]!.click();
		expect(view.state.doc.textContent).toContain('Colour: Blue');
		const date = host.querySelector<HTMLElement>('.dve-sdt[data-sdt-kind="date"]')!;
		expect(click(view, date)).toBe(true);
		const day = document.querySelector<HTMLButtonElement>(
			'.sdt-calendar button[data-date]:not(.outside)',
		)!;
		day.click();
		const stamp = day.dataset.date!;
		expect(
			controlSpans(view.state.doc).find((span) => span.control.kind === 'date')?.control.fullDate,
		).toBe(`${stamp}T00:00:00Z`);
		const xml = await savedXml(loaded, view);
		expect(xml).toContain('<w:t>Blue</w:t>');
		expect(xml).toContain(`w:fullDate="${stamp}T00:00:00Z"`);
		expect(xml.match(/<w:showingPlcHdr\/>/g)?.length).toBe(2);
		view.destroy();
	});

	it('types over a placeholder and honours the lock settings', async () => {
		const { loaded, view } = await open();
		const city = controlSpans(view.state.doc).find((span) => span.control.kind === 'comboBox')!;
		const tr = typeOverPlaceholder(view.state, city.from + 2, 'Lisbon');
		if (!tr || tr === 'blocked') throw new Error('expected an edit');
		view.dispatch(tr);
		expect(view.state.doc.textContent).toContain('City: Lisbon');
		// The notes block is content locked: typing inside it is refused.
		let notesPos = -1;
		view.state.doc.descendants((node, pos) => {
			if (node.type.name === 'paragraph' && node.attrs.contentControl?.alias === 'Notes block')
				notesPos = pos + 2;
		});
		const before = view.state.doc;
		view.dispatch(view.state.tr.insertText('x', notesPos));
		expect(view.state.doc.eq(before)).toBe(true);
		// The name control is sdtLocked: deleting it whole is refused, typing in it is not.
		const name = controlSpans(view.state.doc).find((span) => span.control.kind === 'text')!;
		view.dispatch(view.state.tr.delete(name.from, name.to));
		expect(controlSpans(view.state.doc).some((span) => span.control.kind === 'text')).toBe(true);
		const xml = await savedXml(loaded, view);
		expect(xml).toContain('<w:t>Lisbon</w:t>');
		view.destroy();
	});

	it('edits properties through the control and keeps list values', async () => {
		const { loaded, view } = await open();
		const drop = controlSpans(view.state.doc).find((span) => span.control.kind === 'dropDownList')!;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, drop.from + 1)));
		const items = parseListEntries('Red\nBlue\nViolet', drop.control.items);
		view.dispatch(
			updateControlProperties(view.state, drop, { ...drop.control, alias: 'Shade', items }),
		);
		expect(items.map((item) => item.value)).toEqual(['r', 'b', 'Violet']);
		const xml = await savedXml(loaded, view);
		expect(xml).toContain('<w:alias w:val="Shade"/>');
		expect(xml).toContain('<w:listItem w:displayText="Violet" w:value="Violet"/>');
		expect(xml).not.toContain('w:displayText="Green"');
		view.destroy();
	});

	it('saves a control pasted from another page without what its properties held', async () => {
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">Pick: </w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({
				doc: modelToDoc(loaded.model),
				plugins: [contentControlPlugin(), contentControlViewPlugin()],
			}),
		});
		// The page put a paragraph with a field that fetches a picture where the control's
		// properties go, and a kind no control has.
		const span = document.createElement('span');
		span.dataset.sdt = JSON.stringify({
			id: '4321',
			level: 'inline',
			kind: 'bogus',
			sourceXml: `<w:p xmlns:w="${W}"><w:fldSimple w:instr='INCLUDEPICTURE "http://example.test/a.png"'/></w:p>`,
		});
		span.textContent = 'Pasted';
		view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
		// jsdom has no ClipboardEvent; the paste reads nothing else from the event.
		view.pasteHTML(`<p>${span.outerHTML}</p>`, new Event('paste') as ClipboardEvent);
		expect(controlSpans(view.state.doc).map((item) => item.control.id)).toEqual(['4321']);
		const xml = await savedXml(loaded, view);
		expect(xml).toContain('<w:t>Pasted</w:t>');
		expect(xml).not.toContain('INCLUDEPICTURE');
		expect(xml).not.toContain('w:bogus');
		view.destroy();
		host.remove();
	});
});
