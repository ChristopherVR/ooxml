import { describe, expect, it } from 'vitest';
import { NS, first, parseXml } from '../../xml/index';
import { drawingFillXml } from '../../drawingml/write-fill';
import { columnChart } from './test-support/column-chart';
import {
	chartDrawingColorFromHex,
	chartDrawingGradientFrom,
	chartFillFromAppearance,
	chartLineFromAppearance,
	chartEffectEditsFromAppearance,
	chartThreeDFromAppearance,
	chartTextEditsFromAppearance,
} from './format-values';
import { setChartTargetEffect } from './effects';
import { setChartTargetText } from './text';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';

describe('resolved Office appearance to shared authored values', () => {
	it('converts points/degrees and opacity consistently without silently accepting invalid input', () => {
		expect(chartDrawingColorFromHex('#aabbcc', 40)).toEqual({
			kind: 'srgb',
			value: 'AABBCC',
			transforms: [{ name: 'alpha', value: '60000' }],
		});
		expect(chartDrawingColorFromHex('#aabbcc', 101)).toBeNull();
		expect(
			chartLineFromAppearance({
				kind: 'solid',
				color: '#123456',
				width: 1.5,
				compound: 'dbl',
				join: 'miter',
				beginArrow: 'triangle',
			}),
		).toMatchObject({
			widthEmu: 19050,
			compound: 'dbl',
			join: { kind: 'miter' },
			headEnd: { type: 'triangle' },
		});
		expect(
			chartDrawingGradientFrom({
				type: 'linear',
				stops: [
					{ position: 0, color: '#000000' },
					{ position: 100, color: 'invalid' },
				],
			}),
		).toBeNull();
		expect(() => chartLineFromAppearance({ kind: 'solid', color: '#123456', width: -1 })).toThrow();
	});
	it('converts host-allocated picture fills, preserving requested tiles and opacity', () => {
		const fill = chartFillFromAppearance({
			kind: 'picture',
			picture: { relId: 'rIdImage', tile: true, transparency: 30 },
		})!;
		const root = parseXml(drawingFillXml(fill)!).documentElement;
		expect(first(root, 'tile', NS.a)).toBeDefined();
		expect(first(root, 'stretch', NS.a)).toBeUndefined();
		expect(first(first(root, 'blip', NS.a), 'alphaModFix', NS.a)?.getAttribute('amt')).toBe(
			'70000',
		);
		expect(() =>
			chartFillFromAppearance({ kind: 'picture', picture: { src: 'data:image/png;base64,AA==' } }),
		).toThrow('host');
	});
	it('shares effect conversion with host adapters and retains unrelated effect source', () => {
		let source = columnChart();
		source.spPr = {
			effectsXml: `<a:effectLst xmlns:a="${NS.a}" xmlns:v="urn:v"><v:keep/></a:effectLst>`,
		};
		const edits = chartEffectEditsFromAppearance({
			shadow: { color: '#112233', transparency: 60, blur: 4, distance: 3, angle: 45 },
			glow: { color: '#445566', radius: 5 },
			softEdge: 2,
		});
		for (const edit of edits) source = setChartTargetEffect(source, { kind: 'chartArea' }, edit);
		const saved = writeChartSpace(source);
		expect(saved).toContain('<v:keep');
		expect(saved).toContain('blurRad="50800"');
		expect(saved).toContain('dist="38100"');
		expect(saved).toContain('dir="2700000"');
		expect(parseChartSpace(saved).issues).toEqual([]);
	});
	it('converts one resolved text record into independent pure edits without dropping another effect', () => {
		let source = columnChart();
		const edits = chartTextEditsFromAppearance({
			fill: { kind: 'solid', color: '#123456' },
			rotation: 20,
			effects: {
				shadow: { color: '#000000', blur: 2, distance: 1, angle: 90 },
				glow: { color: '#445566', radius: 3 },
			},
		});
		for (const edit of edits) source = setChartTargetText(source, { kind: 'title' }, edit);
		const saved = writeChartSpace(source);
		expect(saved).toContain('outerShdw');
		expect(saved).toContain('<a:glow');
		expect(saved).toContain('rot="1200000"');
		expect(parseChartSpace(saved).chartSpace.title?.text).toBe('Sales 2026');
	});
	it('converts only supplied 3D values, leaving source camera and other settings untouched', () => {
		expect(
			chartThreeDFromAppearance({
				bevelTop: { type: 'circle', width: 6, height: 3 },
				depth: { size: 4 },
				material: null,
			}),
		).toEqual({
			bevelTop: { preset: 'circle', widthEmu: 76200, heightEmu: 38100 },
			depthEmu: 50800,
			material: null,
		});
	});
});
