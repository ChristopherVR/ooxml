import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { findErrors, withIgnoredError } from '../audit/index';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from './index';
import { miniPackage, ws } from './mini-package-fixtures';

const source = () =>
	miniPackage({
		sheets: [
			{
				name: 'Sheet1',
				xml: ws(
					'<row r="1"><c r="A1" t="str"><v>12</v></c><c r="B1" t="str"><v>13</v></c></row>',
					'',
					'<ignoredErrors><ignoredError sqref="A1 C3:D4" numberStoredAsText="1" formula="1"/></ignoredErrors>',
				),
			},
		],
	});

it('reads ignoredErrors, honours them and writes them back', async () => {
	const wb = await loadXlsx(await source());
	expect(wb.sheets[0]?.ignoredErrors).toEqual([
		{
			ranges: [
				{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
				{ start: { row: 2, col: 2 }, end: { row: 3, col: 3 } },
			],
			flags: ['numberStoredAsText', 'formula'],
		},
	]);
	// A1 is ignored, B1 is not.
	expect(findErrors(wb, 0).map((i) => i.col)).toEqual([1]);
	const session = createEditSession(wb);
	const issue = findErrors(wb, 0)[0]!;
	session.setIgnoredErrors(0, withIgnoredError(wb, issue));
	expect(findErrors(wb, 0)).toEqual([]);
	const saved = await JSZip.loadAsync(await saveXlsx(wb));
	const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string');
	expect(xml).toContain('<ignoredError sqref="A1 C3:D4" numberStoredAsText="1" formula="1"/>');
	expect(xml).toContain('<ignoredError sqref="B1" numberStoredAsText="1"/>');
	const again = await loadXlsx(await saveXlsx(wb));
	expect(again.sheets[0]?.ignoredErrors).toHaveLength(2);
	session.undo();
	expect(findErrors(wb, 0)).toHaveLength(1);
});
