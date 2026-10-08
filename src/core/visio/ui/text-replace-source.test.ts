import { describe, expect, it } from 'vitest';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { VisioPackage } from '../package';
import { cell, fixture, shape, rectangle } from '../test-fixtures';
import {
	visioTextReplaceCommands,
	visioTextReplacePlan,
	type VisioTextReplaceRequest,
} from './text-replace';

const request: VisioTextReplaceRequest = {
	query: 'a',
	replacement: '$&',
	matchCase: true,
	scope: 'all-pages',
	pageId: '0',
	mode: 'all',
};
const local = (id: string, text: string, extra = '') =>
	shape(id, cell('Width', 3) + cell('Height', 1) + rectangle + extra + `<Text>${text}</Text>`);
const source = (extra = '') =>
	fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${local('1', '😀aaa')}${local('2', 'a', extra)}</Shapes>` },
		],
		edit: (zip) => zip.file('custom/opaque.bin', new Uint8Array([0, 255])),
	});

describe('source-backed atomic text replacement plans', () => {
	it('round trips all occurrences while preserving unknown payloads and untouched parts', async () => {
		const bytes = await source('<Unknown payload="preserved"/>'),
			document = await parseVsdx(bytes);
		const plan = visioTextReplacePlan(document, request);
		const result = await editVsdx(bytes, visioTextReplaceCommands(document, plan));
		const actual = await parseVsdx(result.bytes);
		expect(actual.pages[0]!.shapes.map((item) => item.text.plainText)).toEqual(['😀$&$&$&', '$&']);
		const before = await VisioPackage.open(bytes),
			after = await VisioPackage.open(result.bytes);
		for (const path of before.paths())
			if (!result.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		expect(new TextDecoder().decode(await after.readBytes('visio/pages/page1.xml'))).toContain(
			'<Unknown payload="preserved"/>',
		);
	});
	it.each([
		cell('LockTextEdit', 1),
		'<Section N="Field"><Row IX="0"><Cell N="Value" V="a"/></Row></Section>',
		'<Section N="User"><Row N="Dependent"><Cell N="Value" V="1" F="TEXTWIDTH(TheText)"/></Row></Section>',
	])('refuses a late matched protected/field shape atomically %#', async (extra) => {
		const bytes = await source(extra),
			before = bytes.slice(),
			document = await parseVsdx(bytes);
		const plan = visioTextReplacePlan(document, request);
		expect(plan.edits).toHaveLength(2);
		await expect(editVsdx(bytes, visioTextReplaceCommands(document, plan))).rejects.toThrow();
		expect(bytes).toEqual(before);
		expect((await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.plainText).toBe('😀aaa');
	});
	it.each(['<cp IX="0"/>a', '<fld IX="0">a</fld>'])(
		'includes rich matched text and refuses without silent skipping %#',
		async (text) => {
			const bytes = await fixture({
					pages: [{ id: '0', contents: `<Shapes>${local('1', 'a')}${local('2', text)}</Shapes>` }],
				}),
				document = await parseVsdx(bytes);
			const plan = visioTextReplacePlan(document, request);
			expect(plan.replacementCount).toBe(2);
			await expect(editVsdx(bytes, plan.edits)).rejects.toThrow(/Rich text/);
		},
	);
	it('retains no-op targets so rich text is still refused', async () => {
		const bytes = await fixture({
				pages: [{ id: '0', contents: `<Shapes>${local('1', '<cp IX="0"/>a')}</Shapes>` }],
			}),
			document = await parseVsdx(bytes);
		await expect(
			editVsdx(bytes, visioTextReplacePlan(document, { ...request, replacement: 'a' }).edits),
		).rejects.toThrow(/Rich text/);
	});
	it('includes matched master text and refuses the complete source transaction', async () => {
		const bytes = await fixture({
			masters: [{ id: '0', shapes: local('1', 'a') }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${local('1', 'a')}${shape('2', '<Text>a</Text>', 'Master="0" MasterShape="1"')}</Shapes>`,
				},
			],
		});
		const document = await parseVsdx(bytes),
			plan = visioTextReplacePlan(document, request);
		expect(plan.edits).toHaveLength(2);
		await expect(editVsdx(bytes, plan.edits)).rejects.toThrow(/Master-linked/);
	});
	it('deletes literal text and preserves exact repeated no-op source bytes', async () => {
		const bytes = await source(),
			document = await parseVsdx(bytes);
		const noOp = await editVsdx(
			bytes,
			visioTextReplacePlan(document, { ...request, replacement: 'a' }).edits,
		);
		expect(noOp.bytes).toEqual(bytes);
		const saved = await editVsdx(
			bytes,
			visioTextReplacePlan(document, { ...request, replacement: '' }).edits,
		);
		expect(
			(await parseVsdx(saved.bytes)).pages[0]!.shapes.map((item) => item.text.plainText),
		).toEqual(['😀', '']);
	});
});
