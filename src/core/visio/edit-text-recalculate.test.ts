import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { cell, fixture, rectangle, section, shape } from './test-fixtures';

const user = (rows: Record<string, [string, string]>) =>
	section(
		'User',
		Object.entries(rows)
			.map(
				([name, [value, formula]]) =>
					`<Row N="${name}"><Cell N="Value" V="${value}" F="${formula.replaceAll('&', '&amp;').replaceAll('"', '&quot;')}"/></Row>`,
			)
			.join(''),
	);
const replace = (text: string, shapeId = '1') => ({
	type: 'replace-plain-text' as const,
	pageId: '0',
	shapeId,
	text,
});
async function pageXml(bytes: Uint8Array): Promise<string> {
	const pkg = await VisioPackage.open(bytes);
	return new TextDecoder().decode(await pkg.readBytes('visio/pages/page1.xml'));
}
const cacheOf = (xml: string, row: string) =>
	new RegExp(`<Row N="${row}"[^>]*><Cell N="Value" V="([^"]*)"`).exec(xml)?.[1];

const contents = (extra: string, text = 'Old') =>
	`<Shapes>${shape('1', rectangle + `<Text>${text}\n</Text>` + extra)}</Shapes>`;

describe('plain text edits recalculate the cells that read the text', () => {
	it('writes new caches for LEN(SHAPETEXT(TheText)) and everything downstream', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: contents(
						user({
							Len: ['3', 'LEN(SHAPETEXT(TheText))'],
							Double: ['6', 'User.Len*2'],
							Fixed: ['9', '4+5'],
						}),
					),
				},
			],
		});
		const saved = await editVsdx(bytes, [replace('Changed')]);
		const xml = await pageXml(saved.bytes);
		expect(cacheOf(xml, 'Len')).toBe('7');
		expect(cacheOf(xml, 'Double')).toBe('14');
		expect(cacheOf(xml, 'Fixed')).toBe('9');
		expect(xml).toContain('F="LEN(SHAPETEXT(TheText))"');
		expect(saved.diagnostics.map((item) => item.code)).not.toContain('FORMULA_CYCLE');
	});
	it('follows a bare TheText inside text functions and the same text of another shape', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + '<Text>Old\n</Text>')}${shape(
						'2',
						rectangle +
							user({
								Other: ['3', 'LEN(SHAPETEXT(Sheet.1!TheText))'],
								Has: ['1', 'FIND("an",Sheet.1!TheText)'],
							}),
					)}</Shapes>`,
				},
			],
		});
		const xml = await pageXml((await editVsdx(bytes, [replace('Changed')])).bytes);
		expect(cacheOf(xml, 'Other')).toBe('7');
		expect(cacheOf(xml, 'Has')).toBe('3');
	});
	it('drives a size from the text length', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: contents(
						cell('Width', 1.5, 'LEN(SHAPETEXT(TheText))*0.5') + cell('LocPinX', 0.75, 'Width*0.5'),
					),
				},
			],
		});
		const xml = await pageXml((await editVsdx(bytes, [replace('Hello!')])).bytes);
		expect(xml).toContain('<Cell N="Width" V="3"');
		expect(xml).toContain('<Cell N="LocPinX" V="1.5"');
	});
	it('refuses atomically when the dependent formula cannot be evaluated', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: contents(user({ Len: ['3', 'LEN(SHAPETEXT(TheText))+FONT("Arial")'] })),
				},
			],
		});
		const original = bytes.slice();
		await expect(editVsdx(bytes, [replace('Changed')])).rejects.toThrow();
		expect(bytes).toEqual(original);
	});
	it('still refuses a bare TheText in arithmetic and a text-valued cell', async () => {
		for (const formula of ['TheText', 'SHAPETEXT(TheText)', 'SHAPETEXT(TheText)&"x"']) {
			const bytes = await fixture({
				pages: [{ id: '0', contents: contents(user({ Len: ['3', formula] })) }],
			});
			await expect(editVsdx(bytes, [replace('Changed')]), formula).rejects.toMatchObject({
				code: expect.stringMatching(/DEPENDENCY|FORMULA/),
			});
		}
	});
	it('keeps cyclic caches, recalculates the rest and reports the cycle', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: contents(
						user({
							A: ['5', 'User.B+LEN(SHAPETEXT(TheText))'],
							B: ['5', 'User.A'],
							Len: ['3', 'LEN(SHAPETEXT(TheText))'],
						}),
					),
				},
			],
		});
		const saved = await editVsdx(bytes, [replace('Changed')]);
		const xml = await pageXml(saved.bytes);
		expect(cacheOf(xml, 'Len')).toBe('7');
		expect(cacheOf(xml, 'A')).toBe('5');
		expect(cacheOf(xml, 'B')).toBe('5');
		expect(saved.diagnostics.filter((item) => item.code === 'FORMULA_CYCLE')).toHaveLength(2);
	});
});
