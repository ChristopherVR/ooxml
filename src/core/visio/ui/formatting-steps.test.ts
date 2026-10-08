import { describe, expect, it } from 'vitest';
import { parseVsdx } from '../parser';
import { cell, fixture, rectangle, section, shape } from '../test-fixtures';
import { visioTextFontStepCommand } from './formatting';

const page = async (sizes: number[], text: string) =>
	(
		await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape(
							'1',
							cell('Width', 2) +
								cell('Height', 1) +
								rectangle +
								section(
									'Character',
									sizes
										.map((size, index) => `<Row IX="${index}">${cell('Size', size / 72)}</Row>`)
										.join(''),
								) +
								`<Text>${text}</Text>`,
						)}</Shapes>`,
					},
				],
			}),
		)
	).pages[0]!;

describe('rendered font size steps', () => {
	it.each([
		[12, 'increase', 14],
		[12, 'decrease', 11],
		[100, 'increase', 120],
		[100, 'decrease', 72],
		[1, 'decrease', 1],
		[1000, 'increase', 1000],
	] as const)('steps %s points %s to %s', async (size, direction, expected) => {
		expect(visioTextFontStepCommand(await page([size], 'Hello'), '1', direction)?.fontSize).toBe(
			expected,
		);
	});
	it('uses the actual run size when row zero is unused', async () => {
		const source = await page([10, 24], '<cp IX="1"/>Hello');
		expect(source.shapes[0]!.text.fontSize * 72).toBe(10);
		expect(visioTextFontStepCommand(source, '1', 'increase')?.fontSize).toBe(28);
	});
	it('declines mixed run sizes so an increase cannot shrink another run', async () => {
		expect(
			visioTextFontStepCommand(await page([10, 24], '<cp IX="0"/>A<cp IX="1"/>B'), '1', 'increase'),
		).toBeUndefined();
	});
});
