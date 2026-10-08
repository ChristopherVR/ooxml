import { describe, expect, it } from 'vitest';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { VisioPackage } from '../package';
import { cell, fixture, shape } from '../test-fixtures';
import {
	visioTextReplaceCommands,
	visioTextReplacePlan,
	type VisioTextReplaceRequest,
} from './text-replace';

const request: VisioTextReplaceRequest = {
	query: 'cat',
	replacement: 'dog',
	matchCase: true,
	scope: 'current-page',
	pageId: '0',
	mode: 'all',
};
const bytesFor = (text: string, extra = '') =>
	fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', extra + `<Text>${text}</Text>`)}</Shapes>` },
		],
	});
const rows =
	'<Section N="Character"><Row IX="0">' +
	cell('Style', 0) +
	'</Row><Row IX="1">' +
	cell('Style', 1) +
	'</Row><Row IX="2">' +
	cell('Style', 2) +
	'</Row></Section><Section N="Paragraph"><Row IX="0">' +
	cell('HorzAlign', 0) +
	'</Row></Section>';

describe('rich replacement planner dispatch', () => {
	it('round trips cross-run Replace All while keeping source rows and character styles', async () => {
		const bytes = await bytesFor(
				'<cp IX="0"/><pp IX="0"/><tp IX="0"/>A ca<cp IX="1"/>t B cat C<cp IX="2"/> tail\n',
				rows,
			),
			model = await parseVsdx(bytes),
			plan = visioTextReplacePlan(model, request);
		expect(plan.edits[0]!.type).toBe('replace-text-ranges');
		const saved = await editVsdx(bytes, visioTextReplaceCommands(model, plan)),
			text = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.text;
		expect(text.plainText).toBe('A dog B dog C tail');
		expect(
			text.runs.map((run) => ({ text: run.text, bold: run.bold, italic: run.italic })),
		).toEqual([
			{ text: 'A dog', bold: false, italic: false },
			{ text: ' B dog C', bold: true, italic: false },
			{ text: ' tail', bold: false, italic: true },
		]);
		const source = await (await VisioPackage.open(bytes)).readXml('visio/pages/page1.xml'),
			actual = await (await VisioPackage.open(saved.bytes)).readXml('visio/pages/page1.xml');
		expect(
			Array.from(actual.getElementsByTagName('Section')).map((section) => section.toString()),
		).toEqual(
			Array.from(source.getElementsByTagName('Section')).map((section) => section.toString()),
		);
	});
	it.each([
		'',
		'<Section N="Paragraph"><Row IX="01"><Cell N="HorzAlign" V="2"/></Row></Section>',
		'<Section N="Character" Del="1"><Row IX="0"><Cell N="Style" V="2"/></Row></Section>',
	])(
		'keeps admitted no-marker plain source independent of unrelated formatting row proof %#',
		async (extra) => {
			const bytes = await bytesFor('cat\ncat\n', extra),
				model = await parseVsdx(bytes),
				plan = visioTextReplacePlan(model, request);
			const saved = await editVsdx(bytes, visioTextReplaceCommands(model, plan));
			expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.text.plainText).toBe('dog\ndog');
			const noop = visioTextReplacePlan(model, { ...request, replacement: 'cat' });
			expect((await editVsdx(bytes, noop.edits)).bytes).toEqual(bytes);
		},
	);
	it('retains plain paragraph insertion/removal and source-rich refusal', async () => {
		const bytes = await bytesFor('cat\ncat\n'),
			model = await parseVsdx(bytes);
		for (const patch of [{ replacement: 'dog\n' }, { query: 'cat\n', replacement: '' }]) {
			const plan = visioTextReplacePlan(model, { ...request, ...patch });
			expect(plan.edits[0]!.type).toBe('replace-plain-text');
			const result = await editVsdx(bytes, plan.edits);
			expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.text.plainText).toBe(
				patch.query ? 'cat' : 'dog\n\ndog\n',
			);
		}
		const rich = await bytesFor('<cp IX="0"/>cat\n', rows),
			richModel = await parseVsdx(rich);
		await expect(
			editVsdx(rich, visioTextReplacePlan(richModel, { ...request, replacement: 'dog\n' }).edits),
		).rejects.toThrow(/Rich text/);
	});
	it('keeps source authority for stale offsets and protected matched no-ops', async () => {
		const bytes = await bytesFor('cat'),
			model = await parseVsdx(bytes),
			commands = visioTextReplaceCommands(model, visioTextReplacePlan(model, request));
		const changed = await editVsdx(bytes, [
			{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'bat' },
		]);
		await expect(editVsdx(changed.bytes, commands)).rejects.toThrow(/Original text/);
		const locked = await bytesFor('<cp IX="0"/>cat\n', rows + cell('LockTextEdit', 1)),
			lockedModel = await parseVsdx(locked);
		await expect(
			editVsdx(locked, visioTextReplacePlan(lockedModel, { ...request, replacement: 'cat' }).edits),
		).rejects.toThrow();
	});
});
