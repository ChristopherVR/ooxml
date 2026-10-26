import { describe, expect, it } from 'vitest';
import { parseChartSpace, readDisplayNaAsBlank, writeChartSpaceOverSource } from '../index';
import { setChartChrome, setChartProtection, setChartDataTable } from './chrome';

const source =
	'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:protection><c:chartObject val="true"/><c:formatting val="0"/><x:keep xmlns:x="urn:keep" value="retained"/></c:protection><c:chart><c:plotArea><c:dTable><c:showKeys val="1"/><c:extLst><c:ext uri="table"/></c:extLst></c:dTable></c:plotArea><c:extLst><c:ext uri="keep"/></c:extLst></c:chart></c:chartSpace>';

describe('chart chrome edits', () => {
	it('models protection while retaining unknown content and supports explicit removal', () => {
		const parsed = parseChartSpace(source);
		expect(parsed.issues).toEqual([]);
		expect(writeChartSpaceOverSource(source, parsed.chartSpace)).toBe(source);
		const updated = setChartProtection(parsed.chartSpace, { data: true, formatting: null });
		const saved = writeChartSpaceOverSource(source, updated);
		expect(saved).toContain('value="retained"');
		expect(saved).toContain('<c:chartObject val="true"/>');
		expect(saved).toContain('<c:data val="1"/>');
		expect(saved).not.toContain('<c:formatting');
		const cleared = writeChartSpaceOverSource(
			saved,
			setChartProtection(parseChartSpace(saved).chartSpace, null),
		);
		expect(cleared).not.toContain('<c:protection');
	});
	it('patches the display extension and table without replacing unrelated extensions', () => {
		let model = setChartChrome(parseChartSpace(source).chartSpace, {
			displayNaAsBlank: true,
			displayBlanksAs: 'span',
		});
		model = setChartDataTable(model, { showKeys: null, showOutline: true, font: { sizePt: 12 } });
		const saved = writeChartSpaceOverSource(source, model);
		const reopened = parseChartSpace(saved).chartSpace;
		expect(readDisplayNaAsBlank(reopened.chartExtLst)).toBe(true);
		expect(saved).toContain('uri="table"');
		expect(saved).toContain('uri="keep"');
		expect(reopened.plotArea.dataTable?.showKeys).toBeUndefined();
		expect(reopened.plotArea.dataTable?.showOutline).toBe(true);
	});
});
