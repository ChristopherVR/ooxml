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

/** A page shape whose Tabs section is `tabs` and whose text is one tabbed line. */
async function tabbedShape(tabs: string) {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					cell('Width', 4) +
						cell('Height', 1) +
						`<Section N="Tabs"><Row IX="0">${tabs}</Row></Section>` +
						'<Text><tp IX="0"/>a\tb\tc</Text>',
				)}</Shapes>`,
			},
		],
	});
	return (await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.paragraphs![0]!.tabs;
}

describe('tab stops in the model', () => {
	it('reads the stops of the Tabs row a tp marker selects', async () => {
		const tabs = await tabbedShape(
			'<Cell N="Position0" V="1"/><Cell N="Alignment0" V="0"/><Cell N="Position1" V="2.5"/><Cell N="Alignment1" V="2"/>',
		);
		expect(tabs).toEqual([
			{ position: 1, align: 'left' },
			{ position: 2.5, align: 'right' },
		]);
	});

	// Visio 16 stencil masters (ENTITY_M.VSSX, DRILLD_M.VSSX) save their Tabs rows with these
	// cells: numbered from 0, positions in inches and alignments 0 left, 1 centre, 2 right, 3 decimal.
	it('reads the Tabs row Visio 16 saves in the Entity stencil master', async () => {
		const tabs = await tabbedShape(
			'<Cell N="Position0" V="0.25"/><Cell N="Alignment0" V="0"/><Cell N="Position1" V="0.75"/><Cell N="Alignment1" V="0"/>',
		);
		expect(tabs).toEqual([
			{ position: 0.25, align: 'left' },
			{ position: 0.75, align: 'left' },
		]);
	});

	it('reads the single right stop Visio 16 saves in the Pivot Node master', async () => {
		expect(await tabbedShape('<Cell N="Position0" V="1"/><Cell N="Alignment0" V="2"/>')).toEqual([
			{ position: 1, align: 'right' },
		]);
	});

	// The Tabs row a page shape holds after Visio 16 re-saved it (positions carry U="IN").
	it('reads the centre and decimal stops of a page re-saved by Visio 16', async () => {
		const tabs = await tabbedShape(
			"<Cell N='Position0' V='3.25' U='IN'/><Cell N='Alignment0' V='1'/><Cell N='Position1' V='4.5' U='IN'/><Cell N='Alignment1' V='3'/>",
		);
		expect(tabs).toEqual([
			{ position: 3.25, align: 'center' },
			{ position: 4.5, align: 'decimal' },
		]);
	});
});
