import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import { mailMergeField } from './authoring';
import { findLabel, labelCatalogue, labelVendors, searchLabels } from './label-catalogue';
import { createLabels, labelPresets, updateLabels } from './labels';

describe('label catalogue', () => {
	it('fits every product on its sheet, once per vendor', () => {
		expect(labelCatalogue.length).toBeGreaterThanOrEqual(35);
		const seen = new Set<string>();
		for (const spec of labelCatalogue) {
			const key = `${spec.vendor}:${spec.product}`;
			expect(seen.has(key), key).toBe(false);
			seen.add(key);
			const across = spec.columns * spec.width + (spec.columns - 1) * spec.gap;
			const down = spec.rows * spec.height + (spec.rows - 1) * spec.verticalGap;
			// One twip of rounding per label at most.
			expect(spec.left + across, key).toBeLessThanOrEqual(spec.pageWidth + spec.columns);
			expect(spec.top + down, key).toBeLessThanOrEqual(spec.pageHeight + spec.rows);
		}
		expect(labelVendors).toEqual([
			'Avery US Letter',
			'Online Labels',
			'Avery A4',
			'Avery Zweckform',
			'HERMA A4',
		]);
	});

	it('keeps the three earlier presets with their measurements', () => {
		expect(labelPresets.map(({ product }) => product)).toEqual(['5160', '5163', '4464']);
		expect(findLabel('5160')).toMatchObject({
			columns: 3,
			rows: 10,
			left: 270,
			top: 720,
			width: 3780,
			height: 1440,
			gap: 180,
			verticalGap: 0,
		});
		expect(findLabel('5163')).toMatchObject({ left: 225, gap: 270, width: 5760, height: 2880 });
		expect(findLabel('4464', 'HERMA A4')).toMatchObject({ left: 0, top: 28, columns: 3, rows: 8 });
		expect(findLabel('L7160')).toMatchObject({ top: 859, left: 409, columns: 3, rows: 7 });
	});

	it('finds products by number, alias or description within a vendor', () => {
		expect(findLabel('8160')?.product).toBe('5160');
		expect(findLabel('j 8163')?.product).toBe('L7163');
		expect(searchLabels('816', 'Avery US Letter').map(({ product }) => product)).toEqual([
			'5160',
			'5161',
			'5162',
			'5163',
			'5164',
			'5165',
			'5167',
			'5168',
		]);
		expect(searchLabels('', 'HERMA A4').map(({ product }) => product)).toEqual(['4464']);
		expect(searchLabels('return', 'Avery US Letter').map(({ product }) => product)).toEqual([
			'5167',
			'5195',
		]);
		expect(searchLabels('nothing like this')).toEqual([]);
	});

	it('builds sheets for any product and spacer rows for labels with a vertical gap', () => {
		const sheet = createLabels({ product: 'L7651' });
		const grid = sheet.blocks[0];
		if (grid?.type !== 'table') throw new Error('Missing label table');
		expect(grid.rows).toHaveLength(13);
		expect(grid.rows[0]).toHaveLength(9);
		const base = findLabel('5160')!;
		const gapped = createLabels({
			product: 'custom',
			spec: { ...base, rows: 3, verticalGap: base.gap },
		});
		const table = gapped.blocks[0];
		if (table?.type !== 'table') throw new Error('Missing label table');
		expect(table.rowProperties?.map(({ heightTwips }) => heightTwips)).toEqual([
			1440, 180, 1440, 180, 1440,
		]);
		const first = documentParagraphs(gapped.blocks)[0]!;
		first.runs = [mailMergeField('Name')];
		const updated = updateLabels({ ...gapped, mailMerge: { mainDocumentType: 'mailingLabels' } });
		const rows = (updated.blocks[0] as typeof table).rows;
		const filled = (row: number) =>
			rows[row]!.filter((cell) =>
				cell.paragraphs.some((p) => p.runs.some((r) => r.field || r.fieldCode)),
			).length;
		expect([0, 1, 2, 3, 4].map(filled)).toEqual([3, 0, 3, 0, 3]);
		expect(() => createLabels({ product: 'x', spec: { ...base, rows: 0 } })).toThrow(RangeError);
	});
});
