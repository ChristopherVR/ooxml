import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { loadDocx } from './parse';
import type { Paragraph } from './model';
import { fieldMarkerNodeSpec } from './ui/break-note-schema';
import { fieldGuardPlugin, fieldsBalanced } from './ui/field-guard';
import { markSpecs } from './ui/schema-marks';
import { inlineNodeRun, runToInlineNodes } from './ui/run-adapter';

async function source(result: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>L</w:t></w:r><w:fldSimple w:instr=" REF Target ">${result}</w:fldSimple><w:r><w:t>R</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});

it.each(['<w:r><w:t/></w:r>', '<w:r/>', '<w:r><w:rPr><w:b/></w:rPr></w:r>'])(
	'preserves an explicitly empty simple result without placeholder text: %s',
	async (result) => {
		const loaded = await loadDocx(await source(result));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs.map((run) => run.text).join('')).toBe('LR');
		expect(paragraph.runs.filter((run) => run.fieldChar).map((run) => run.fieldChar)).toEqual([
			'begin',
			'separate',
			'end',
		]);
		expect(paragraph.runs.find((run) => run.fieldCode)?.fieldCode).toBe(' REF Target ');
		const bytes = await loaded.save();
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).not.toContain('[Field]');
		const reloaded = await loadDocx(bytes);
		const runs = (reloaded.model.blocks[0] as Paragraph).runs;
		expect(runs.map((run) => run.text).join('')).toBe('LR');
		expect(runs.find((run) => run.fieldCode)?.fieldCode).toBe(' REF Target ');
		if (result.includes('<w:b/>')) expect(runs.some((run) => run.bold)).toBe(true);
	},
);

it('retains the existing display fallback when a simple field has no saved result runs', async () => {
	const loaded = await loadDocx(await source(''));
	expect((loaded.model.blocks[0] as Paragraph).runs[1]).toMatchObject({
		text: '[Field]',
		field: { instr: 'REF Target', simple: true },
	});
});

it('keeps a supported nontext field cache instead of replacing it with a placeholder', async () => {
	const loaded = await loadDocx(await source('<w:r><w:br w:type="page"/></w:r>'));
	expect((loaded.model.blocks[0] as Paragraph).runs[1]).toMatchObject({
		text: '',
		break: 'page',
		field: { instr: 'REF Target', simple: true },
	});
});

it('preserves an empty field through editor insertion, export and local undo', async () => {
	const loaded = await loadDocx(await source('<w:r><w:t/></w:r>'));
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const doc = schema.node(
		'doc',
		null,
		schema.node(
			'paragraph',
			null,
			paragraph.runs.flatMap((run) => runToInlineNodes(run, schema)),
		),
	);
	expect(doc.textContent).toBe('LR');
	expect(fieldsBalanced(doc)).toBe(true);
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, 5),
		plugins: [history(), fieldGuardPlugin()],
	});
	state = state.applyTransaction(state.tr.insertText('Result')).state;
	expect(state.doc.textContent).toBe('LResultR');
	const exportState = async () => {
		const runs = Array.from({ length: state.doc.firstChild!.childCount }, (_, index) =>
			inlineNodeRun(state.doc.firstChild!.child(index))!,
		);
		return loadDocx(await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }));
	};
	const edited = (await exportState()).model.blocks[0] as Paragraph;
	expect(edited.runs.filter((run) => run.field).map((run) => run.text)).toEqual(['Result']);
	expect(edited.runs.find((run) => run.fieldCode)?.fieldCode).toBe(' REF Target ');
	undo(state, (tr) => {
		state = state.applyTransaction(tr).state;
	});
	expect(state.doc.textContent).toBe('LR');
	expect(fieldsBalanced(state.doc)).toBe(true);
	const restored = (await exportState()).model.blocks[0] as Paragraph;
	expect(restored.runs.map((run) => run.text).join('')).toBe('LR');
	expect(restored.runs.find((run) => run.fieldCode)?.fieldCode).toBe(' REF Target ');
});
