import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { columnChart } from './test-support/column-chart';
import { applyChartAppearance } from './apply-appearance';

describe('resolved appearance differences', () => {
	it('keeps theme fills and arrows while changing outline width and opacity', () => {
		const source = columnChart();
		source.spPr = {
			fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } },
			line: {
				fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2', transforms: [] } },
				widthEmu: 12700,
				tailEnd: { type: 'arrow', width: 'lg' },
			},
		};
		const before = {
			fill: { kind: 'solid' as const, color: '#123456', transparency: 0 },
			line: { kind: 'solid' as const, color: '#ABCDEF', width: 1 },
		};
		const next = applyChartAppearance(source, { kind: 'chartArea' }, before, {
			fill: { ...before.fill, transparency: 25 },
			line: { ...before.line, width: 3 },
		});
		const parsed = parseChartSpace(writeChartSpace(next)).chartSpace;
		expect(parsed.spPr?.fill).toMatchObject({
			kind: 'solid',
			color: { kind: 'scheme', value: 'accent1', transforms: [{ name: 'alpha', value: '75000' }] },
		});
		expect(parsed.spPr?.line).toMatchObject({
			widthEmu: 38100,
			fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
			tailEnd: { type: 'arrow', width: 'lg' },
		});
		expect(source.spPr.line!.widthEmu).toBe(12700);
	});
	it('keeps a theme shadow colour and unknown effects while changing blur', () => {
		const source = columnChart();
		source.spPr = {
			effectsXml:
				'<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:outerShdw blurRad="12700" dist="12700" dir="0"><a:schemeClr val="dk1"/></a:outerShdw><a:extension/></a:effectLst>',
		};
		const shadow = { color: '#000000', blur: 1, distance: 1, angle: 0 };
		const next = applyChartAppearance(
			source,
			{ kind: 'chartArea' },
			{ effects: { shadow } },
			{ effects: { shadow: { ...shadow, blur: 3 } } },
		);
		const xml = writeChartSpace(next);
		expect(xml).toContain('blurRad="38100"');
		expect(xml).toContain('<a:schemeClr val="dk1"/>');
		expect(xml).toContain('<a:extension/>');
	});
	it('preserves gradient stop theme tokens while editing geometry', () => {
		const source = columnChart();
		source.spPr = {
			fill: {
				kind: 'gradient',
				angle: 0,
				stops: [
					{ position: 0, color: { kind: 'scheme', value: 'accent1', transforms: [] } },
					{ position: 100, color: { kind: 'scheme', value: 'accent2', transforms: [] } },
				],
			},
		};
		const gradient = {
			type: 'linear' as const,
			angle: 0,
			stops: [
				{ position: 0, color: '#123456' },
				{ position: 100, color: '#ABCDEF' },
			],
		};
		const next = applyChartAppearance(
			source,
			{ kind: 'chartArea' },
			{ fill: { kind: 'gradient', gradient } },
			{ fill: { kind: 'gradient', gradient: { ...gradient, angle: 45 } } },
		);
		expect(writeChartSpace(next)).toContain('<a:schemeClr val="accent1"');
		expect(next.spPr!.fill).toMatchObject({ angle: 45 });
	});
});
