import { describe, expect, it } from 'vitest';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { cell, fixture, rectangle, shape } from '../test-fixtures';
import { visioChangeCaseCommand, visioChangeCaseShape } from './change-case';

const characters =
	'<Section N="Character"><Row IX="0">' +
	cell('Style', 0) +
	'</Row><Row IX="1">' +
	cell('Style', 2) +
	'</Row></Section>';
const rich = '<cp IX="0"/><pp IX="0"/>hello <cp IX="1"/>rich WOrld\n<cp IX="0"/>second line\n';
const source = (text = rich) =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 3) + cell('Height', 1) + rectangle + characters + `<Text>${text}</Text>`)}${shape('2', cell('Width', 1) + cell('Height', 1) + rectangle)}</Shapes>`,
			},
		],
	});
const italics = (document: Awaited<ReturnType<typeof parseVsdx>>) =>
	document.pages[0]!.shapes[0]!.text.runs.map((run) => [run.text, run.italic]);

describe('Change Case range commands', () => {
	for (const [mode, expected] of [
		['upper', 'HELLO RICH WORLD\nSECOND LINE'],
		['lower', 'hello rich world\nsecond line'],
		['sentence', 'Hello rich world\nSecond line'],
		['capitalize', 'Hello Rich World\nSecond Line'],
		['toggle', 'HELLO RICH woRLD\nSECOND LINE'],
	] as const)
		it(`applies ${mode} without flattening runs`, async () => {
			const bytes = await source();
			const before = await parseVsdx(bytes);
			const command = visioChangeCaseCommand(before.pages[0]!, '1', mode)!;
			const saved = await editVsdx(bytes, [command]);
			const after = await parseVsdx(saved.bytes);
			expect(after.pages[0]!.shapes[0]!.text.plainText).toBe(expected);
			const runs = italics(after);
			expect(runs.map(([, italic]) => italic)).toEqual(italics(before).map(([, italic]) => italic));
			expect(runs.map(([text]) => (text as string).length)).toEqual(
				italics(before).map(([text]) => (text as string).length),
			);
		});

	it('returns nothing when the case already matches or the shape has no text', async () => {
		const document = await parseVsdx(await source('ABC'));
		const page = document.pages[0]!;
		expect(visioChangeCaseCommand(page, '1', 'upper')).toBeUndefined();
		expect(visioChangeCaseCommand(page, '1', 'lower')).toMatchObject({
			type: 'replace-text-ranges',
			ranges: [{ start: 0, end: 3, text: 'abc' }],
		});
		expect(visioChangeCaseShape(page, '2')).toBeUndefined();
		expect(visioChangeCaseCommand(page, '2', 'upper')).toBeUndefined();
		expect(visioChangeCaseCommand(page, '1', 'bogus' as never)).toBeUndefined();
	});
});
