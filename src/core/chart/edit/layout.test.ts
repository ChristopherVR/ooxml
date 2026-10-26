import { describe, expect, it } from 'vitest';
import { parseChartSpace, writeChartSpace } from '../index';
import { setChartLayout } from './layout';
import { setChartLegendEntries } from './legend-entries';
import { setChartPointLabel } from './point-label';
import { editChartLine } from './line-options';
import { drawingLineXml } from '../../drawingml/write-line';

const source =
	'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout><c:manualLayout><c:x val="0.1"/><c:extLst><c:ext uri="manual"/></c:extLst></c:manualLayout><c:extLst><c:ext uri="outer"/></c:extLst></c:layout></c:plotArea><c:legend><c:legendEntry><c:idx val="4"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr lang="ja-JP"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:legendEntry></c:legend></c:chart></c:chartSpace>';

describe('shared layout and legend entry edits', () => {
	it('preserves producer attributes while editing an existing manual coordinate', () => {
		const xml = source.replace(
			'<c:x val="0.1"/>',
			'<c:x xmlns:x="urn:producer" val="0.1" x:keep="original"/>',
		);
		const saved = writeChartSpace(
			setChartLayout(parseChartSpace(xml).chartSpace, 'plotArea', { x: 0.5 }),
		);
		expect(saved).toContain('val="0.5"');
		expect(saved).toContain('x:keep="original"');
		expect(saved).toContain('urn:producer');
	});
	it('switches exclusive dash choices without resurrecting the old choice on Automatic reset', () => {
		const preset = { dash: 'dash', widthEmu: 12700 };
		const custom = editChartLine(preset, { customDash: [{ dash: 200000, space: 100000 }] });
		expect(drawingLineXml(custom)).toContain('<a:custDash>');
		expect(drawingLineXml(custom)).not.toContain('prstDash');
		const next = editChartLine(custom, { dash: 'dot' });
		expect(drawingLineXml(next)).toContain('val="dot"');
		expect(drawingLineXml(next)).not.toContain('custDash');
		const reset = editChartLine(next, { dash: null });
		expect(drawingLineXml(reset)).not.toMatch(/custDash|prstDash/);
		expect(() => editChartLine(preset, { dash: 'dot', customDash: [] })).toThrow('either');
	});
	it('preserves point-label source extensions and language while editing its text and outline', () => {
		const input = parseChartSpace(
			source.replace(
				'<c:plotArea>',
				'<c:plotArea><c:barChart><c:ser><c:idx val="9"/><c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB"/><a:t>Value</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></c:spPr><c:extLst><c:ext uri="keep"/></c:extLst></c:dLbl></c:dLbls></c:ser></c:barChart>',
			),
		).chartSpace;
		const saved = writeChartSpace(
			setChartPointLabel(input, 0, 0, { text: 'Values', shape: { line: { widthEmu: 25400 } } }),
		);
		expect(saved).toContain('uri="keep"');
		expect(saved).toContain('lang="en-GB"');
		expect(saved).toContain('val="accent2"');
		expect(saved).toContain('w="25400"');
		expect(
			parseChartSpace(saved).chartSpace.plotArea.groups[0]!.series[0]!.dataLabels!.labels[0]!.tx!
				.text,
		).toBe('Values');
	});
	it('keeps both extension lists after manual placement and automatic reset', () => {
		const input = parseChartSpace(source).chartSpace;
		const moved = setChartLayout(input, 'plotArea', { x: 0.25, y: 0.4 });
		const saved = writeChartSpace(moved);
		expect(saved).toContain('<c:x val="0.25"/>');
		expect(saved).toContain('uri="manual"');
		expect(saved).toContain('uri="outer"');
		const reset = writeChartSpace(
			setChartLayout(parseChartSpace(saved).chartSpace, 'plotArea', null),
		);
		expect(reset).not.toContain('<c:x ');
		expect(reset).toContain('uri="manual"');
		expect(reset).toContain('uri="outer"');
		expect(input.plotArea.layout?.x).toBe(0.1);
	});
	it('patches an entry font without pinning its theme or erasing language', () => {
		const next = setChartLegendEntries(parseChartSpace(source).chartSpace, [
			{ index: 4, deleted: false, text: { sizePt: 16 } },
		]);
		const saved = writeChartSpace(next);
		expect(saved).toContain('lang="ja-JP"');
		expect(saved).toContain('<a:schemeClr val="accent1"/>');
		expect(saved).toContain('sz="1600"');
		expect(() => setChartLegendEntries(next, [{ index: 4 }, { index: 4 }])).toThrow(/unique/);
	});
});
