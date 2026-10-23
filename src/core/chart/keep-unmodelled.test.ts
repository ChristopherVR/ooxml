import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import { parseChartSpace } from './parse-space';
import { writeChartSpace } from './write-space';
import { setChartTitle } from './edit/title';
import { withUnmodelledChartChildren } from './keep-unmodelled';

const source =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
	'<c:roundedCorners val="0"/>' +
	'<c:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
	'<c:pivotSource><c:name>[Book1.xlsx]Sheet1!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource>' +
	'<c:protection><c:chartObject val="1"/></c:protection>' +
	'<c:chart><c:autoTitleDeleted val="1"/><c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts>' +
	'<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
	'<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
	'<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
	'<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
	'</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

describe('withUnmodelledChartChildren', () => {
	it('carries pivot, protection and colour-map children into an edited part, in schema order', () => {
		const { chartSpace, issues } = parseChartSpace(source);
		expect(issues.map((issue) => issue.message).join(' ')).toContain('pivotSource');
		const written = writeChartSpace(setChartTitle(chartSpace, 'Pivot sales'), {
			declarationBreak: '\n',
		});
		expect(written).not.toContain('pivotSource');
		const kept = withUnmodelledChartChildren(source, written);
		const order = [
			'c:roundedCorners ',
			'c:clrMapOvr ',
			'c:pivotSource>',
			'c:protection>',
			'c:chart>',
		];
		const positions = order.map((name) => kept.indexOf(`<${name}`));
		expect(positions.every((at) => at >= 0)).toBe(true);
		expect([...positions].sort((a, b) => a - b)).toEqual(positions);
		const title = kept.indexOf('<c:title>');
		const pivotFmts = kept.indexOf('<c:pivotFmts>');
		expect(title).toBeLessThan(kept.indexOf('<c:autoTitleDeleted'));
		expect(kept.indexOf('<c:autoTitleDeleted')).toBeLessThan(pivotFmts);
		expect(pivotFmts).toBeLessThan(kept.indexOf('<c:plotArea>'));
		expect(kept.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')).toBe(true);
		expect(parseChartSpace(kept).chartSpace.title?.text).toBe('Pivot sales');
	});

	it('returns the written part itself when the source has nothing unmodelled', () => {
		const plain = source
			.replace(/<c:clrMapOvr[^>]*\/>/, '')
			.replace(/<c:pivotSource>.*<\/c:pivotSource>/, '')
			.replace(/<c:protection>.*<\/c:protection>/, '')
			.replace(/<c:pivotFmts>.*<\/c:pivotFmts>/, '');
		const written = writeChartSpace(parseChartSpace(plain).chartSpace);
		expect(withUnmodelledChartChildren(plain, written)).toBe(written);
	});
});
