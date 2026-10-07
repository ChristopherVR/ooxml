import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { Fragment, Schema, Slice } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { loadDocx, type Paragraph } from './index';
import { parseOnOff } from './simple-types';
import { getW, parseXml, WORD_NS } from './xml';
import { fieldMarkerNodeSpec } from './ui/break-note-schema';
import { markSpecs } from './ui/schema-marks';
import { inlineNodeRun, runToInlineNodes } from './ui/run-adapter';
import { replaceSimpleFieldResult, simpleFieldPasteSlice } from './ui/simple-field-input';
import { fieldClipboardSlice } from './ui/field-clipboard';
import { parseBlocksFromContainer } from './block-parser';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});

async function loadField(content: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p>${content}<w:r><w:t>tail</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const simple = (flags: string, result = 'ABC') =>
	`<w:fldSimple w:instr=" QUOTE ABC " ${flags}><w:r><w:t>${result}</w:t></w:r></w:fldSimple>`;

it.each([
	['w:fldLock="true" w:dirty="0"', { locked: true, dirty: false }],
	['w:fldLock="off" w:dirty="on"', { locked: false, dirty: true }],
	['', undefined],
] as const)('models simple flags independently from absence (%s)', async (flags, expected) => {
	const loaded = await loadField(simple(flags));
	const run = (loaded.model.blocks[0] as Paragraph).runs[0]!;
	expect(run).toMatchObject({ field: { instr: 'QUOTE ABC', simple: true } });
	expect(run.fieldFlags).toEqual(expected);
	const node = runToInlineNodes(run, schema)[0]!;
	expect(inlineNodeRun(node)?.fieldFlags).toEqual(expected);
});

it('retains simple flags when formatting splits its result and rebuilds the wrapper', async () => {
	const loaded = await loadField(simple('w:fldLock="false" w:dirty="true"'));
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const [run, tail] = paragraph.runs;
	const runs = [{ ...run!, text: 'A', bold: true }, { ...run!, text: 'BC', italic: true }, tail!];
	const saved = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
	const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
	const wrapper = parseXml(xml).getElementsByTagNameNS(WORD_NS, 'fldSimple')[0]!;
	expect(parseOnOff(getW(wrapper, 'fldLock'))).toBe(false);
	expect(parseOnOff(getW(wrapper, 'dirty'))).toBe(true);
	const reloaded = await loadDocx(saved);
	expect((reloaded.model.blocks[0] as Paragraph).runs.slice(0, 2)).toMatchObject([
		{ fieldFlags: { locked: false, dirty: true } },
		{ fieldFlags: { locked: false, dirty: true } },
	]);
});

it('preserves flags on each complex marker after editing code and cached text', async () => {
	const loaded = await loadField(
		'<w:r><w:fldChar w:fldCharType="begin" w:fldLock="true" w:dirty="false"/></w:r><w:r><w:instrText> QUOTE ABC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate" w:fldLock="0"/></w:r><w:r><w:t>ABC</w:t></w:r><w:r><w:fldChar w:fldCharType="end" w:dirty="1"/></w:r>',
	);
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const runs = paragraph.runs.map((run) => {
		const node = runToInlineNodes(run, schema)[0]!;
		return {
			...inlineNodeRun(node)!,
			...(run.fieldCode !== undefined ? { fieldCode: ' QUOTE XYZ ' } : {}),
			...(run.text === 'ABC' ? { text: 'XYZ' } : {}),
		};
	});
	expect(runs[0]).toMatchObject({ fieldFlags: { locked: true, dirty: false } });
	expect(runs[2]).toMatchObject({ fieldFlags: { locked: false } });
	expect(runs[4]).toMatchObject({ fieldFlags: { dirty: true } });
	const reloaded = await loadDocx(
		await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
	);
	const next = (reloaded.model.blocks[0] as Paragraph).runs;
	expect(next[0]).toMatchObject({ fieldFlags: { locked: true, dirty: false } });
	expect(next[2]).toMatchObject({ fieldFlags: { locked: false } });
	expect(next[4]).toMatchObject({ fieldFlags: { dirty: true } });
	expect(next.find((run) => run.text === 'XYZ')).toMatchObject({ field: { instr: 'QUOTE XYZ' } });
});

it('places empty simple-cache flags on its converted begin marker', async () => {
	const loaded = await loadField(simple('w:fldLock="false" w:dirty="0"', ''));
	const paragraph = loaded.model.blocks[0] as Paragraph;
	expect(paragraph.runs[0]).toMatchObject({
		fieldChar: 'begin',
		fieldFlags: { locked: false, dirty: false },
	});
	expect(paragraph.runs.slice(1).some((run) => run.fieldFlags)).toBe(false);
	const reloaded = await loadDocx(
		await loaded.save({
			...loaded.model,
			blocks: [{ ...paragraph, runs: [{ text: 'before' }, ...paragraph.runs] }],
		}),
	);
	expect((reloaded.model.blocks[0] as Paragraph).runs[1]).toMatchObject({
		fieldChar: 'begin',
		fieldFlags: { locked: false, dirty: false },
	});
});

it.each(['X', ''])(
	'retains structural flags during typed result replacement (%s)',
	async (replacement) => {
		const loaded = await loadField(simple('w:fldLock="true" w:dirty="false"'));
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
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1, 4) });
		const tr = replaceSimpleFieldResult(state, 1, 4, replacement)!;
		const runs = Array.from({ length: tr.doc.firstChild!.childCount }, (_, index) =>
			inlineNodeRun(tr.doc.firstChild!.child(index))!,
		);
		expect(runs[0]).toMatchObject({ fieldFlags: { locked: true, dirty: false } });
		const reloaded = await loadDocx(
			await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
		);
		expect((reloaded.model.blocks[0] as Paragraph).runs[0]).toMatchObject({
			fieldFlags: { locked: true, dirty: false },
		});
		if (replacement) {
			const copied = fieldClipboardSlice(tr.doc.slice(1, 2));
			expect(inlineNodeRun(copied.content.firstChild!)).not.toHaveProperty('fieldFlags');
		}
	},
);

it('retains simple wrapper flags through contextual paste but strips literal result metadata', async () => {
	const loaded = await loadField(simple('w:fldLock="false" w:dirty="true"'));
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
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, 2, 3) });
	const slice = simpleFieldPasteSlice(
		new Slice(Fragment.from(schema.text('X', [schema.marks.bold!.create()])), 0, 0),
		state,
	);
	expect(inlineNodeRun(slice.content.firstChild!)).toMatchObject({
		fieldFlags: { locked: false, dirty: true },
		bold: true,
	});
	const copied = fieldClipboardSlice(slice);
	expect(inlineNodeRun(copied.content.firstChild!)).toEqual({ text: 'X', bold: true });
});

it('isolates nested and cross-paragraph cached lock state while dirty stays on markers', () => {
	const begin = (lock: string, instr: string) =>
		`<w:r><w:fldChar w:fldCharType="begin" w:fldLock="${lock}" w:dirty="true"/></w:r><w:r><w:instrText>${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
	const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
	const body = parseXml(
		`<w:body xmlns:w="${WORD_NS}"><w:p>${begin('true', 'QUOTE')}<w:r><w:t>before</w:t></w:r></w:p><w:p>${begin('false', 'PAGE')}<w:r><w:t>1</w:t></w:r></w:p><w:p>${end}<w:r><w:t>after</w:t></w:r>${end}<w:r><w:t>outside</w:t></w:r></w:p></w:body>`,
	).documentElement!;
	const runs = parseBlocksFromContainer(body)
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.filter((run) => run.text);
	expect(runs).toEqual([
		{ text: 'before', field: { instr: 'QUOTE' }, fieldFlags: { locked: true } },
		{ text: '1', field: { instr: 'PAGE' }, fieldFlags: { locked: false } },
		{ text: 'after', field: { instr: 'QUOTE' }, fieldFlags: { locked: true } },
		{ text: 'outside' },
	]);
	const header = parseXml(
		`<w:hdr xmlns:w="${WORD_NS}"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>`,
	).documentElement!;
	expect((parseBlocksFromContainer(header)[0] as Paragraph).runs).toEqual([{ text: 'header' }]);
});

for (const input of ['type', 'paste'] as const)
	for (const identified of [false, true])
		it.each([undefined, { locked: false, dirty: false }])(
			`${input} uses only target structural flags (${identified ? 'identified' : 'anonymous'} target, %j)`,
			(flags) => {
				const target = runToInlineNodes(
					{
						text: 'AB',
						field: { instr: 'QUOTE target', simple: true },
						...(identified && { fieldInstanceId: 'target' }),
						...(flags && { fieldFlags: flags }),
					},
					schema,
				);
				const adjacent = runToInlineNodes(
					{
						text: 'CD',
						field: { instr: 'QUOTE adjacent', simple: true },
						fieldInstanceId: 'adjacent',
						fieldFlags: { locked: true, dirty: true },
					},
					schema,
				);
				const doc = schema.node(
					'doc',
					null,
					schema.node('paragraph', null, [...target, ...adjacent]),
				);
				const stale = schema.marks.runProperties!.create({
					props: {
						fieldFlags: { locked: true, dirty: true },
						fieldInstanceId: 'source',
						caps: true,
					},
				});
				let state = EditorState.create({
					doc,
					selection: TextSelection.create(doc, input === 'type' ? 2 : 1, input === 'type' ? 2 : 3),
				});
				let tr;
				if (input === 'type') {
					state = state.apply(state.tr.setStoredMarks([stale]));
					tr = replaceSimpleFieldResult(state, 1, 3, 'X')!;
				} else {
					const slice = simpleFieldPasteSlice(
						new Slice(Fragment.from(schema.text('X', [stale])), 0, 0),
						state,
					);
					tr = state.tr.replaceSelection(slice);
				}
				const result = inlineNodeRun(tr.doc.nodeAt(1)!)!;
				expect(result.fieldFlags).toEqual(flags);
				expect(result.fieldInstanceId).toBe(identified ? 'target' : undefined);
				expect(result).toMatchObject({
					text: 'X',
					caps: true,
					field: { instr: 'QUOTE target', simple: true },
				});
				expect(inlineNodeRun(tr.doc.nodeAt(2)!)).toMatchObject({
					text: 'CD',
					fieldInstanceId: 'adjacent',
					fieldFlags: { locked: true, dirty: true },
				});
			},
		);
