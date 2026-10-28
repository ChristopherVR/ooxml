import { expect, it } from 'vitest';
import { parseRegionValueColorPositions } from './value-color-positions';
import { resolveRegionValueColors } from './value-colors';
import { resolveValueColorStops } from './data';
const native =
	'<cx:valueColorPositions count="3"><cx:min><cx:number val="10"/></cx:min><cx:mid><cx:percent val="25"/></cx:mid><cx:max><cx:number val="90"/></cx:max></cx:valueColorPositions>';
it('resolves schema min/mid/max numeric and percentage breakpoints with themed colors', () => {
	const options = resolveRegionValueColors(
		{
			valueColorsXml:
				'<cx:valueColors><cx:minColor><a:srgbClr val="FF0000"/></cx:minColor><cx:midColor><a:schemeClr val="accent2"/></cx:midColor><cx:maxColor><a:srgbClr val="0000FF"/></cx:maxColor></cx:valueColors>',
			valueColorPositionsXml: native,
		},
		(color) => (color ? (color.kind === 'scheme' ? '#00FF00' : '#' + color.value) : undefined),
	);
	expect(options.valueColors).toEqual(['#FF0000', '#00FF00', '#0000FF']);
	expect(resolveValueColorStops(options.valueColors, options.valueColorPositions, 0, 100)).toEqual([
		{ color: '#FF0000', position: 0.1 },
		{ color: '#00FF00', position: 0.25 },
		{ color: '#0000FF', position: 0.9 },
	]);
});
it('retains default extremes and missing midpoint without shifting color alignment', () => {
	expect(
		parseRegionValueColorPositions(
			'<cx:valueColorPositions count="3"><cx:max><cx:extremeValue/></cx:max></cx:valueColorPositions>',
		),
	).toEqual([{ kind: 'min' }, { kind: 'percent', value: 50 }, { kind: 'max' }]);
	expect(
		parseRegionValueColorPositions(
			'<cx:valueColorPositions><cx:min><cx:extremeValue/></cx:min><cx:max><cx:percent val="85"/></cx:max></cx:valueColorPositions>',
		),
	).toEqual([{ kind: 'min' }, { kind: 'percent', value: 85 }]);
});
it('keeps explicit pos shorthand compatibility and declines invalid numeric values', () => {
	expect(
		parseRegionValueColorPositions(
			'<cx:valueColorPositions count="2"><cx:pos type="num" val="20"/><cx:pos type="max"/></cx:valueColorPositions>',
		),
	).toEqual([{ kind: 'number', value: 20 }, { kind: 'max' }]);
	expect(
		parseRegionValueColorPositions(
			'<cx:valueColorPositions><cx:pos type="min"/><cx:pos type="num" val="20"/><cx:pos type="percent" val="75"/><cx:pos type="max"/></cx:valueColorPositions>',
		),
	).toEqual([
		{ kind: 'min' },
		{ kind: 'number', value: 20 },
		{ kind: 'percent', value: 75 },
		{ kind: 'max' },
	]);
	expect(parseRegionValueColorPositions(native.replace('val="10"', 'val="Infinity"'))?.[0]).toEqual(
		{ kind: 'min' },
	);
});
