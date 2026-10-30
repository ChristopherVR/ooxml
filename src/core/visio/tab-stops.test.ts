import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { VISIO_DEFAULT_TAB_SPACING, visioTabStarts } from './tab-stops';
import { cell, fixture, shape } from './test-fixtures';

const seg = (width: number, decimal = width) => ({ width, decimal });

describe('tab stop positions', () => {
	it('moves to the next explicit stop and then to the default spacing', () => {
		const tabs = [{ position: 1, align: 'left' as const }];
		expect(visioTabStarts([seg(0.3), seg(0.2), seg(0.2)], tabs, 0, 0)).toEqual([0, 1, 1.5]);
		expect(VISIO_DEFAULT_TAB_SPACING).toBe(0.5);
		// No stops: every tab goes to the next multiple of the default spacing.
		expect(visioTabStarts([seg(0.2), seg(0.2)], [], 0.1, 0.1)).toEqual([0.1, 0.6]);
	});

	it('aligns centre, right and decimal stops on the stop', () => {
		const starts = (align: 'center' | 'right' | 'decimal') =>
			visioTabStarts([seg(0.1), seg(0.4, 0.3)], [{ position: 2, align }], 0, 0)[1];
		expect(starts('center')).toBeCloseTo(1.8);
		expect(starts('right')).toBeCloseTo(1.6);
		expect(starts('decimal')).toBeCloseTo(1.7);
	});

	it('never starts text before the text already placed', () => {
		const tabs = [{ position: 0.2, align: 'right' as const }];
		expect(visioTabStarts([seg(0.1), seg(0.6)], tabs, 0, 0)[1]).toBeCloseTo(0.1);
	});
});

describe('tab stops in the model', () => {
	it('reads the stops of the Tabs row a tp marker selects', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						cell('Width', 4) +
							cell('Height', 1) +
							'<Section N="Tabs"><Row IX="0"><Cell N="Position1" V="1"/><Cell N="Alignment1" V="0"/><Cell N="Position2" V="2.5"/><Cell N="Alignment2" V="2"/></Row></Section>' +
							'<Text><tp IX="0"/>a\tb\tc</Text>',
					)}</Shapes>`,
				},
			],
		});
		const text = (await parseVsdx(bytes)).pages[0]!.shapes[0]!.text;
		expect(text.paragraphs![0]!.tabs).toEqual([
			{ position: 1, align: 'left' },
			{ position: 2.5, align: 'right' },
		]);
	});
});
