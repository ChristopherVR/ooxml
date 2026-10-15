import { describe, expect, it } from 'vitest';
import { type CarriedRefEdits, patchCarriedPart } from './carried-refs';

const edits: CarriedRefEdits = { shifts: [], deleted: [], renamed: [['Data', 'Renamed']] };

describe('carried part scans stay linear', () => {
	it('handles many unclosed formula tags quickly', () => {
		const xml = `<c:chartSpace>${'<c:f '.repeat(50_000)}`;
		const started = Date.now();
		patchCarriedPart('xl/charts/chart1.xml', xml, edits);
		expect(Date.now() - started).toBeLessThan(2000);
	});

	it('handles many unclosed worksheetSource tags quickly', () => {
		const xml = `<pivotCacheDefinition>${'<worksheetSource a="1" '.repeat(50_000)}`;
		const started = Date.now();
		patchCarriedPart('xl/pivotCache/pivotCacheDefinition1.xml', xml, edits);
		expect(Date.now() - started).toBeLessThan(2000);
	});

	it('still rewrites chart formulas and a worksheet source', () => {
		const chart =
			'<c:chartSpace><c:f>Data!$A$1</c:f><c:f x="1">Other!$A$1</c:f><c:f>Data!$B$1</c:f></c:chartSpace>';
		expect(patchCarriedPart('xl/charts/chart1.xml', chart, edits)).toBe(
			'<c:chartSpace><c:f>Renamed!$A$1</c:f><c:f x="1">Other!$A$1</c:f><c:f>Renamed!$B$1</c:f></c:chartSpace>',
		);
		const pivot =
			'<pivotCacheDefinition><worksheetSource ref="A1:C7" sheet="Data"/></pivotCacheDefinition>';
		expect(patchCarriedPart('xl/pivotCache/pivotCacheDefinition1.xml', pivot, edits)).toContain(
			'sheet="Renamed"',
		);
	});
});
