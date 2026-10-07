import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, shape } from './test-fixtures';

it('preserves independent foreground and background alpha without a colored layer', async () => {
	const document = await parseVsdx(
		await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + cell('FillPattern', 26) + cell('FillForegnd', '#00ff00') + cell('FillBkgnd', '#0000ff') + cell('FillForegndTrans', 0.25) + cell('FillBkgndTrans', 0.6))}</Shapes>`,
				},
			],
		}),
	);
	const style = document.pages[0]!.shapes[0]!.style;
	expect(style.fillOpacity).toBe(1);
	expect(style.fillGradient?.stops).toEqual([
		{ offset: 0, color: '#0000ff', opacity: 0.4 },
		{ offset: 0.5, color: '#00ff00', opacity: 0.75 },
		{ offset: 1, color: '#0000ff', opacity: 0.4 },
	]);
	expect(document.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(false);
});

it('keeps unsupported pattern and incomplete modern gradient diagnostics', async () => {
	for (const settings of [
		cell('FillPattern', 31),
		cell('FillPattern', 25) + cell('FillGradientEnabled', 1),
	]) {
		const document = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', rectangle + settings + cell('FillForegnd', '#00ff00') + cell('FillBkgnd', '#0000ff'))}</Shapes>`,
					},
				],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.style.fillGradient).toBeUndefined();
		expect(document.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(
			true,
		);
	}
});
