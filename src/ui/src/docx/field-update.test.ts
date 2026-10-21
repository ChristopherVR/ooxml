// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { history } from 'prosemirror-history';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { seqInstruction } from './caption-commands';
import { insertCaption } from './caption-commands';
import { runToInlineNodes, inlineNodeRun } from 'ooxml-core/docx/ui';
import {
	createDocument,
	loadDocx,
	type DocumentModel,
	type Paragraph,
	type TextRun,
} from 'ooxml-core/docx';
import { readFile } from 'node:fs/promises';
import { URL as NodeURL } from 'node:url';
import { fieldLockFixture } from '../../../core/docx/test-support/field-lock-fixture';
import {
	complexField,
	paragraph as styleParagraph,
	simpleField,
	styleRefFixture,
} from '../../../core/docx/test-support/style-ref-fixture';
import { updateFields } from './field-update';
import { docToModel } from './model-adapter';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => (document.body.innerHTML = ''));

const field = (text: string, instr: string) =>
	schema.text(text, [schema.marks.field!.create({ instr, simple: true })]);
const flaggedField = (text: string, instr: string, locked: boolean, simple = true) =>
	runToInlineNodes(
		{
			text,
			field: { instr, ...(simple && { simple: true }) },
			fieldFlags: { locked, dirty: true },
		},
		schema,
	)[0]!;
const paragraph = (id: string, content: unknown[], bookmarks: string[] = []) =>
	schema.nodes.paragraph!.create({ id, bookmarks }, content as never);

function editor(...blocks: ReturnType<typeof paragraph>[]) {
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, {
		state: EditorState.create({
			doc: schema.node('doc', null, blocks),
			schema,
			plugins: [history()],
		}),
	});
}
const texts = (view: EditorView) => {
	const out: string[] = [];
	view.state.doc.forEach((node) => out.push(node.textContent));
	return out;
};
const pages = (id: string) => ({ p0: '1', p1: '4' })[id];
const native = JSON.parse(
	await readFile(
		new NodeURL(
			'../../../core/docx/__fixtures__/field-locks/native-reference.json',
			import.meta.url,
		),
		'utf8',
	),
) as {
	cases: {
		kind: 'simple' | 'complex';
		method: string;
		before: string[];
		after: string[];
		locked: boolean[];
	}[];
};

describe('updateFields', () => {
	it('does not merge an adjacent anonymous unlocked reference into its locked peer', () => {
		const view = editor(
			paragraph('p0', [schema.text('New title')], ['Target']),
			paragraph('p1', [
				flaggedField('Old title', 'REF Target', false),
				flaggedField('Locked title', 'REF Target', true),
			]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)[1]).toBe('New titleLocked title');
		view.destroy();
	});
	it.each(native.cases)(
		'matches native $kind $method locked field results and survives save/reload',
		async (reference) => {
			const loaded = await loadDocx(await fieldLockFixture(reference.kind));
			const blocks = loaded.model.blocks as Paragraph[];
			const view = editor(
				...blocks.map((block, index) =>
					paragraph(
						block.id,
						block.runs.flatMap((run) => runToInlineNodes(run, schema)),
						index === 0 ? ['Target'] : [],
					),
				),
			);
			expect(texts(view).slice(1)).toEqual(reference.before);
			expect(updateFields(view, () => '1')).toBe(true);
			expect(texts(view).slice(1)).toEqual(reference.after);
			const updated = blocks.map((block, index) => {
				const runs: Paragraph['runs'] = [];
				view.state.doc.child(index).forEach((node) => runs.push(inlineNodeRun(node)!));
				return { ...block, runs };
			});
			const reloaded = await loadDocx(await loaded.save({ ...loaded.model, blocks: updated }));
			const results = (reloaded.model.blocks as Paragraph[])
				.slice(1)
				.map((block) => block.runs.find((run) => run.field)!);
			expect(results.map((run) => run.text)).toEqual(reference.after);
			expect(results.map((run) => run.fieldFlags?.locked)).toEqual(reference.locked);
			view.destroy();
		},
	);
	it.each([true, false])(
		'preserves locked simple=%s references while updating unlocked peers',
		(simple) => {
			const view = editor(
				paragraph('p0', [schema.text('New title')], ['_Ref2']),
				paragraph('p1', [
					flaggedField('Locked title', 'REF _Ref2', true, simple),
					schema.text('|'),
					flaggedField('9', 'PAGEREF _Ref2', true, simple),
					schema.text('|'),
					flaggedField('Old title', 'REF _Ref2', false, simple),
					schema.text('|'),
					flaggedField('9', 'PAGEREF _Ref2', false, simple),
				]),
			);
			expect(updateFields(view, pages)).toBe(true);
			expect(texts(view)[1]).toBe('Locked title|9|New title|1');
			expect(inlineNodeRun(view.state.doc.child(1).firstChild!)?.fieldFlags).toEqual({
				locked: true,
				dirty: true,
			});
			undo(view.state, view.dispatch);
			expect(texts(view)[1]).toBe('Locked title|9|Old title|9');
			view.destroy();
		},
	);

	it('preserves a locked SEQ cache during manual refresh and automatic caption insertion', () => {
		const view = editor(
			paragraph('p0', [schema.text('Figure '), flaggedField('7', seqInstruction('Figure'), true)]),
		);
		expect(updateFields(view, pages)).toBe(false);
		expect(texts(view)).toEqual(['Figure 7']);
		expect(
			insertCaption(view, { label: 'Figure', labelText: 'Figure', text: 'New', position: 'below' }),
		).toBe(true);
		// The existing ordinal policy counts locked fields, but never replaces their saved result.
		expect(texts(view)).toEqual(['Figure 7', 'Figure 2: New']);
		undo(view.state, view.dispatch);
		expect(texts(view)).toEqual(['Figure 7']);
		view.destroy();
	});

	it('updates adjacent identical references independently and retains their field identities', () => {
		const instr = 'REF _Ref2';
		const result = (id: string) =>
			schema.text('Old', [
				schema.marks.field!.create({ instr, simple: true }),
				schema.marks.runProperties!.create({ props: { fieldInstanceId: id } }),
			]);
		const view = editor(
			paragraph('p0', [schema.text('New title')], ['_Ref2']),
			paragraph('p1', [result('first'), result('second')]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)[1]).toBe('New titleNew title');
		const ids: string[] = [];
		view.state.doc
			.child(1)
			.forEach((node) =>
				ids.push(
					node.marks.find((mark) => mark.type.name === 'runProperties')!.attrs.props
						.fieldInstanceId,
				),
			);
		expect(ids).toEqual(['first', 'second']);
		undo(view.state, view.dispatch);
		expect(texts(view)[1]).toBe('OldOld');
	});
	it('renumbers captions, then refreshes REF and PAGEREF results', () => {
		const view = editor(
			paragraph('p0', [schema.text('Intro ')], []),
			paragraph(
				'p1',
				[schema.text('Figure '), field('7', seqInstruction('Figure')), schema.text(': Chart')],
				['_Ref1'],
			),
			paragraph('p2', [
				schema.text('See '),
				field('Figure 9', ' REF _Ref1 \\h '),
				schema.text(' on page '),
				field('9', ' PAGEREF _Ref1 \\h '),
			]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)).toEqual(['Intro ', 'Figure 1: Chart', 'See Figure 1 on page 4']);
	});

	it('follows a heading text REF, keeps unsupported switches and missing targets, and undoes in one step', () => {
		const view = editor(
			paragraph('p0', [schema.text('New title')], ['_Ref2']),
			paragraph('p1', [
				field('Old title', ' REF _Ref2 \\h '),
				schema.text('|'),
				field('x', ' REF _Ref2 \\n \\h '),
				schema.text('|'),
				field('gone', ' REF _Missing \\h '),
			]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)[1]).toBe('New title|x|gone');
		expect(updateFields(view, pages)).toBe(false);
		undo(view.state, view.dispatch);
		expect(texts(view)[1]).toBe('Old title|x|gone');
	});

	it('gives a body STYLEREF the nearest text in its style before it, else after it', () => {
		const heading = (id: string, text: string) =>
			schema.nodes.paragraph!.create({ id, style: 'Heading1' }, [schema.text(text)]);
		const view = editor(
			heading('p0', 'Scope'),
			paragraph('p1', [schema.text('See '), field('old', ' STYLEREF 1 ')]),
			paragraph('p2', [
				field('old', ' STYLEREF "Doc Number" '),
				schema.text('|'),
				field('x', ' STYLEREF 1 \\w '),
			]),
			schema.nodes.paragraph!.create({ id: 'p3', style: 'DocNumber' }, [schema.text('DOC-5')]),
		);
		const paragraphOf = (id: string, style: string | undefined, text: string): Paragraph => ({
			type: 'paragraph',
			id,
			...(style ? { style } : {}),
			runs: [{ text }],
		});
		const model: DocumentModel = {
			...createDocument(),
			blocks: [
				paragraphOf('p0', 'Heading1', 'Scope'),
				paragraphOf('p1', undefined, 'See old'),
				paragraphOf('p2', undefined, 'old|x'),
				paragraphOf('p3', 'DocNumber', 'DOC-5'),
			],
			paragraphStyles: {
				docDefaults: {},
				styles: {
					Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} },
					DocNumber: { id: 'DocNumber', name: 'Doc Number', formatting: {} },
				},
				warnings: [],
			},
		};
		expect(updateFields(view, pages)).toBe(false);
		expect(updateFields(view, pages, model)).toBe(true);
		expect(texts(view)).toEqual(['Scope', 'See Scope', 'DOC-5|x', 'DOC-5']);
	});

	it('reads a body STYLEREF from its own paragraph and from captions renumbered with it', () => {
		const styled = (text: string) => runToInlineNodes({ text, style: 'EntryName' }, schema)[0]!;
		const view = editor(
			schema.nodes.paragraph!.create({ id: 'p0', style: 'Caption' }, [
				schema.text('Figure '),
				field('9', seqInstruction('Figure')),
			]),
			paragraph('p1', [styled('Alder')]),
			paragraph('p2', [
				styled('Ash'),
				schema.text(', see '),
				field('old', ' STYLEREF "Entry Name" '),
				schema.text(' in '),
				field('old', ' STYLEREF Caption '),
			]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { Caption: { id: 'Caption', name: 'caption', formatting: {} } },
				warnings: [],
			},
			characterStyles: {
				docDefaults: {},
				styles: {
					EntryName: { id: 'EntryName', type: 'character', name: 'Entry Name', formatting: {} },
				},
				warnings: [],
			},
		};
		expect(updateFields(view, pages, docToModel(view.state.doc, prior))).toBe(true);
		expect(texts(view)).toEqual(['Figure 1', 'Alder', 'Ash, see Ash in Figure 1']);
	});

	it('formats a CHARFORMAT STYLEREF result like the first letter of its field name', () => {
		const complex = (instr: string, result: string, code: Partial<TextRun>) =>
			(
				[
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ' ' },
					{ ...code, text: '', fieldCode: `${instr} ` },
					{ text: '', fieldChar: 'separate' },
					{ text: result, bold: true, field: { instr } },
					{ text: '', fieldChar: 'end' },
				] as TextRun[]
			).flatMap((run) => runToInlineNodes(run, schema));
		const view = editor(
			schema.nodes.paragraph!.create({ id: 'p0', style: 'DocNumber' }, [schema.text('DOC-5')]),
			paragraph('p1', [
				...complex('STYLEREF "Doc Number" \\* CHARFORMAT', 'old', { italic: true }),
				schema.text('|'),
				...complex('STYLEREF "Doc Number" \\* MERGEFORMAT', 'old', { italic: true }),
			]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { DocNumber: { id: 'DocNumber', name: 'Doc Number', formatting: {} } },
				warnings: [],
			},
		};
		expect(updateFields(view, pages, docToModel(view.state.doc, prior))).toBe(true);
		expect(texts(view)).toEqual(['DOC-5', 'DOC-5|DOC-5']);
		const results: TextRun[] = [];
		view.state.doc.child(1).forEach((child) => {
			const run = child.isText ? inlineNodeRun(child) : undefined;
			if (run?.field) results.push(run);
		});
		expect(results.map((run) => [run.text, run.italic, run.bold])).toEqual([
			['DOC-5', true, undefined],
			['DOC-5', undefined, true],
		]);
		expect(updateFields(view, pages, docToModel(view.state.doc, prior))).toBe(false);
	});

	it('brings a CHARFORMAT result up to date when only its formatting differs', () => {
		const view = editor(
			schema.nodes.paragraph!.create({ id: 'p0', style: 'DocNumber' }, [schema.text('DOC-5')]),
			paragraph(
				'p1',
				(
					[
						{ text: '', fieldChar: 'begin' },
						{ text: '', italic: true, fieldCode: ' STYLEREF "Doc Number" \\* CHARFORMAT ' },
						{ text: '', fieldChar: 'separate' },
						{ text: 'DOC-5', bold: true, field: { instr: 'STYLEREF "Doc Number" \\* CHARFORMAT' } },
						{ text: '', fieldChar: 'end' },
					] as TextRun[]
				).flatMap((run) => runToInlineNodes(run, schema)),
			),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { DocNumber: { id: 'DocNumber', name: 'Doc Number', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		expect(update()).toBe(true);
		const result: TextRun[] = [];
		view.state.doc.child(1).forEach((child) => {
			const run = child.isText ? inlineNodeRun(child) : undefined;
			if (run?.field) result.push(run);
		});
		expect(result.map((run) => [run.text, run.italic, run.bold])).toEqual([
			['DOC-5', true, undefined],
		]);
		expect(update()).toBe(false);
	});

	it('brings a CHARFORMAT result with line breaks up to date once', () => {
		for (const heading of ['\nDOC-5', '\n']) {
			const view = editor(
				schema.nodes.paragraph!.create(
					{ id: 'p0', style: 'DocNumber' },
					runToInlineNodes({ text: heading }, schema),
				),
				paragraph(
					'p1',
					(
						[
							{ text: '', fieldChar: 'begin' },
							{ text: '', italic: true, fieldCode: ' STYLEREF "Doc Number" \\* CHARFORMAT ' },
							{ text: '', fieldChar: 'separate' },
							{
								text: heading,
								bold: true,
								field: { instr: 'STYLEREF "Doc Number" \\* CHARFORMAT' },
							},
							{ text: '', fieldChar: 'end' },
						] as TextRun[]
					).flatMap((run) => runToInlineNodes(run, schema)),
				),
			);
			const prior: DocumentModel = {
				...createDocument(),
				paragraphStyles: {
					docDefaults: {},
					styles: { DocNumber: { id: 'DocNumber', name: 'Doc Number', formatting: {} } },
					warnings: [],
				},
			};
			const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
			expect(update(), heading).toBe(true);
			const result: [string, boolean | undefined, boolean | undefined][] = [];
			view.state.doc.child(1).forEach((child) => {
				const run =
					child.isText || child.type.name === 'hardBreak' ? inlineNodeRun(child) : undefined;
				if (run?.field) result.push([run.text, run.italic, run.bold]);
			});
			expect(result, heading).toEqual(
				heading === '\n'
					? [['\n', true, undefined]]
					: [
							['\n', true, undefined],
							['DOC-5', true, undefined],
						],
			);
			expect(update(), heading).toBe(false);
			view.destroy();
		}
	});

	it('clears a body STYLEREF whose paragraph in the style is empty, and fills it again', async () => {
		const loaded = await loadDocx(
			await styleRefFixture({
				body:
					styleParagraph('Scope', 'Heading1') +
					styleParagraph(
						'',
						undefined,
						`${complexField('STYLEREF 1', ['Scope'])}<w:r><w:t xml:space="preserve"> | </w:t></w:r>${simpleField('STYLEREF 1', 'Scope')}`,
					),
			}),
		);
		const blocks = loaded.model.blocks as Paragraph[];
		const view = editor(
			...blocks.map((block) =>
				schema.nodes.paragraph!.create(
					{ id: block.id, ...(block.style ? { style: block.style } : {}) },
					block.runs.flatMap((run) => runToInlineNodes(run, schema)),
				),
			),
		);
		const update = () => updateFields(view, pages, docToModel(view.state.doc, loaded.model));
		view.dispatch(view.state.tr.delete(1, 1 + 'Scope'.length));
		expect(update()).toBe(true);
		expect(texts(view)).toEqual(['', ' | ']);
		// Both fields keep their instruction, saved as fields with an empty result.
		const saved = await loadDocx(await loaded.save(docToModel(view.state.doc, loaded.model)));
		const runs = (saved.model.blocks[1] as Paragraph).runs;
		expect(
			runs.filter((run) => run.fieldCode !== undefined).map((run) => run.fieldCode!.trim()),
		).toEqual(['STYLEREF 1', 'STYLEREF 1']);
		expect(runs.filter((run) => run.field).map((run) => run.text)).toEqual([]);
		expect(runs.map((run) => run.text).join('')).toBe(' | ');
		// Typed again, the heading fills both empty results.
		view.dispatch(view.state.tr.insertText('Method', 1));
		expect(update()).toBe(true);
		expect(texts(view)).toEqual(['Method', 'Method | Method']);
		expect(update()).toBe(false);
	});

	it('keeps a line break in a body STYLEREF result a line break, written once', () => {
		const complex = (instr: string, result: string) =>
			(
				[
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ` ${instr} ` },
					{ text: '', fieldChar: 'separate' },
					{ text: result, field: { instr } },
					{ text: '', fieldChar: 'end' },
				] as TextRun[]
			).flatMap((run) => runToInlineNodes(run, schema));
		const view = editor(
			schema.nodes.paragraph!.create(
				{ id: 'p0', style: 'Heading1' },
				runToInlineNodes({ text: 'Chapter 1\nIntroduction' }, schema),
			),
			paragraph('p1', [
				...complex('STYLEREF 1', 'Chapter 1\nIntroduction'),
				schema.text('|'),
				...runToInlineNodes(
					{ text: 'Chapter 1\nIntroduction', field: { instr: ' STYLEREF 1 ', simple: true } },
					schema,
				),
				schema.text('|'),
				...complex('STYLEREF 1', 'old'),
			]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		// Text as it shows, a line break as `<br>`; a raw `\n` in a text node would show as `\n`.
		const shown = () => {
			let out = '';
			view.state.doc.child(1).forEach((child) => {
				if (child.isText) out += child.text!;
				else if (child.type.name === 'hardBreak') out += '<br>';
			});
			return out;
		};
		expect(update()).toBe(true);
		expect(shown()).toBe(
			'Chapter 1<br>Introduction|Chapter 1<br>Introduction|Chapter 1<br>Introduction',
		);
		expect(update()).toBe(false);
		// The breaks are the field's, saved as breaks in its result.
		const runs = docToModel(view.state.doc, prior).blocks[1] as Paragraph;
		expect(runs.runs.filter((run) => run.field).map((run) => run.text)).toEqual([
			'Chapter 1\nIntroduction',
			'Chapter 1\nIntroduction',
			'Chapter 1\nIntroduction',
		]);
		// An empty heading clears each result whole; the simple field keeps its instruction.
		view.dispatch(view.state.tr.delete(1, view.state.doc.child(0).nodeSize - 1));
		expect(update()).toBe(true);
		expect(shown()).toBe('||');
		const cleared = (docToModel(view.state.doc, prior).blocks[1] as Paragraph).runs;
		expect(cleared.filter((run) => run.field)).toEqual([]);
		expect(
			cleared.filter((run) => run.fieldCode !== undefined).map((run) => run.fieldCode!.trim()),
		).toEqual(['STYLEREF 1', 'STYLEREF 1', 'STYLEREF 1']);
	});

	it('keeps a line break that starts a body STYLEREF result, written once', () => {
		const complex = (result: string) =>
			(
				[
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ' STYLEREF 1 ' },
					{ text: '', fieldChar: 'separate' },
					{ text: result, field: { instr: 'STYLEREF 1' } },
					{ text: '', fieldChar: 'end' },
				] as TextRun[]
			).flatMap((run) => runToInlineNodes(run, schema));
		const view = editor(
			schema.nodes.paragraph!.create(
				{ id: 'p0', style: 'Heading1' },
				runToInlineNodes({ text: '\nScope' }, schema),
			),
			paragraph('p1', [
				...complex('\nScope'),
				schema.text('|'),
				...complex('old'),
				schema.text('|'),
				...runToInlineNodes(
					{ text: '\nScope', field: { instr: ' STYLEREF 1 ', simple: true } },
					schema,
				),
			]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		const shown = () => {
			let out = '';
			view.state.doc.child(1).forEach((child) => {
				if (child.isText) out += child.text!;
				else if (child.type.name === 'hardBreak') out += '<br>';
			});
			return out;
		};
		expect(update()).toBe(true);
		expect(shown()).toBe('<br>Scope|<br>Scope|<br>Scope');
		expect(update()).toBe(false);
		// Cleared, each result goes whole; the simple field keeps its instruction.
		view.dispatch(view.state.tr.delete(1, view.state.doc.child(0).nodeSize - 1));
		expect(update()).toBe(true);
		expect(shown()).toBe('||');
		const cleared = (docToModel(view.state.doc, prior).blocks[1] as Paragraph).runs;
		expect(
			cleared.filter((run) => run.fieldCode !== undefined).map((run) => run.fieldCode!.trim()),
		).toEqual(['STYLEREF 1', 'STYLEREF 1', 'STYLEREF 1']);
	});

	it('finds a field whose result is only a line break, and updates it again', () => {
		const complex = (result: string) =>
			(
				[
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ' STYLEREF 1 ' },
					{ text: '', fieldChar: 'separate' },
					{ text: result, field: { instr: 'STYLEREF 1' } },
					{ text: '', fieldChar: 'end' },
				] as TextRun[]
			).flatMap((run) => runToInlineNodes(run, schema));
		const view = editor(
			schema.nodes.paragraph!.create(
				{ id: 'p0', style: 'Heading1' },
				runToInlineNodes({ text: '\n' }, schema),
			),
			paragraph('p1', [
				...complex('old'),
				schema.text('|'),
				...runToInlineNodes(
					{ text: 'old', field: { instr: ' STYLEREF 1 ', simple: true } },
					schema,
				),
			]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		const shown = () => {
			let out = '';
			view.state.doc.child(1).forEach((child) => {
				if (child.isText) out += child.text!;
				else if (child.type.name === 'hardBreak') out += '<br>';
			});
			return out;
		};
		expect(update()).toBe(true);
		expect(shown()).toBe('<br>|<br>');
		expect(update()).toBe(false);
		// The heading gets its text again: both fields follow.
		view.dispatch(view.state.tr.insertText('Scope', 1));
		expect(update()).toBe(true);
		expect(shown()).toBe('Scope<br>|Scope<br>');
		expect(update()).toBe(false);
		// Back to a line break, then empty: each field keeps its instruction.
		view.dispatch(view.state.tr.delete(1, 1 + 'Scope'.length));
		expect(update()).toBe(true);
		view.dispatch(view.state.tr.delete(1, view.state.doc.child(0).nodeSize - 1));
		expect(update()).toBe(true);
		expect(shown()).toBe('|');
		const cleared = (docToModel(view.state.doc, prior).blocks[1] as Paragraph).runs;
		expect(
			cleared.filter((run) => run.fieldCode !== undefined).map((run) => run.fieldCode!.trim()),
		).toEqual(['STYLEREF 1', 'STYLEREF 1']);
	});

	it('writes a body STYLEREF result into the text a revision keeps, never into deleted text', () => {
		const instr = 'STYLEREF "Doc Number"';
		const result = (deleted: string, inserted?: string) =>
			(
				[
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ` ${instr} ` },
					{ text: '', fieldChar: 'separate' },
					{
						text: deleted,
						field: { instr },
						revision: { kind: 'delete', id: '1', author: 'A' },
					},
					...(inserted === undefined
						? []
						: [
								{
									text: inserted,
									field: { instr },
									revision: { kind: 'insert', id: '2', author: 'A' },
								},
							]),
					{ text: '', fieldChar: 'end' },
				] as TextRun[]
			).flatMap((run) => runToInlineNodes(run, schema));
		const view = editor(
			schema.nodes.paragraph!.create({ id: 'p0', style: 'DocNumber' }, [schema.text('DOC-3')]),
			paragraph('p1', result('DOC-1', 'DOC-2')),
			paragraph('p2', result('DOC-1')),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { DocNumber: { id: 'DocNumber', name: 'Doc Number', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		const results = (index: number) => {
			const out: [string, string | undefined][] = [];
			view.state.doc.child(index).forEach((child) => {
				const run = child.isText ? inlineNodeRun(child) : undefined;
				if (run?.field) out.push([run.text, run.revision?.kind]);
			});
			return out;
		};
		expect(update()).toBe(true);
		expect(results(1)).toEqual([
			['DOC-1', 'delete'],
			['DOC-3', 'insert'],
		]);
		// A result a revision deletes whole stays as it is until the revision is settled.
		expect(results(2)).toEqual([['DOC-1', 'delete']]);
		expect(update()).toBe(false);
	});

	it('keeps a line break in a REF result a line break, written once', () => {
		const view = editor(
			paragraph('p0', runToInlineNodes({ text: 'Line one\nLine two' }, schema), ['Target']),
			paragraph(
				'p1',
				runToInlineNodes(
					{ text: 'Line one\nLine two', field: { instr: 'REF Target', simple: true } },
					schema,
				),
			),
			paragraph('p2', [field('old', 'REF Target')]),
		);
		const shown = (index: number) => {
			let out = '';
			view.state.doc.child(index).forEach((child) => {
				if (child.isText) out += child.text!;
				else if (child.type.name === 'hardBreak') out += '<br>';
			});
			return out;
		};
		expect(updateFields(view, pages)).toBe(true);
		expect([shown(1), shown(2)]).toEqual(['Line one<br>Line two', 'Line one<br>Line two']);
		expect(updateFields(view, pages)).toBe(false);
	});

	it('brings fields that read other fields up to date in one update', () => {
		const view = editor(
			schema.nodes.paragraph!.create({ id: 'p0', style: 'Heading1' }, [
				schema.text('Chapter '),
				field('Old', 'REF Title'),
			]),
			paragraph('p1', [schema.text('New name')], ['Title']),
			paragraph('p2', [field('Chapter Old', ' STYLEREF 1 ')]),
			paragraph('p3', [schema.text('See '), field('Chapter Old', ' STYLEREF 1 ')], ['Quote']),
			paragraph('p4', [field('See Chapter Old', 'REF Quote')]),
		);
		const prior: DocumentModel = {
			...createDocument(),
			paragraphStyles: {
				docDefaults: {},
				styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
				warnings: [],
			},
		};
		const update = () => updateFields(view, pages, docToModel(view.state.doc, prior));
		expect(update()).toBe(true);
		expect(texts(view)).toEqual([
			'Chapter New name',
			'New name',
			'Chapter New name',
			'See Chapter New name',
			'See Chapter New name',
		]);
		expect(update()).toBe(false);
	});

	it('does nothing in a read-only view', () => {
		const view = editor(paragraph('p0', [field('1', seqInstruction('Figure'))]));
		view.setProps({ editable: () => false });
		expect(updateFields(view, pages)).toBe(false);
	});
});
