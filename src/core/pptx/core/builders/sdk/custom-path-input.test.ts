/**
 * Custom geometry on new shapes: several paths with their own fill / stroke
 * flags, arcs written as cubic Beziers. Each saved case is loaded back, so
 * the written `a:custGeom` and the reloaded paths agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { ShapePptxElement } from '../../types';
import { createShapeElement } from './ElementFactory';
import type { ShapeOptions } from './types';

async function roundTrip(
	options: ShapeOptions,
): Promise<{ custGeom: string; spPr: string; shape: ShapePptxElement | undefined }> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addShape('rect', options).build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	return {
		custGeom: /<a:custGeom>[\s\S]*<\/a:custGeom>/u.exec(xml)?.[0] ?? '',
		spPr: /<p:spPr>[\s\S]*?<\/p:spPr>/u.exec(xml)?.[0] ?? '',
		shape: loaded.slides[0]?.elements[0] as ShapePptxElement | undefined,
	};
}

describe('custom geometry input', () => {
	it('writes one a:path per input with its own flags, arcs as cubics, and reads them back', async () => {
		const { custGeom, shape } = await roundTrip({
			width: 100,
			height: 50,
			fill: { type: 'solid', color: '#FF0000' },
			paths: [
				{ d: 'M0 0 H100 V50 H0 Z', width: 100, height: 50 },
				{ d: 'M10 25 a15 15 0 0 1 30 0', width: 100, height: 50, fill: 'none', stroke: false },
			],
		});
		const paths = custGeom.match(/<a:path\b[\s\S]*?<\/a:path>/gu) ?? [];
		expect(paths).toHaveLength(2);
		// The 100 x 50 px shape: paths in EMU of its size.
		expect(paths[0]).toMatch(/<a:path w="952500" h="476250"/u);
		expect(paths[0]).not.toContain('fill=');
		expect(paths[1]).toContain('fill="none"');
		expect(paths[1]).toContain('stroke="0"');
		expect(paths[1]).toContain('<a:cubicBezTo>');
		expect(paths[1]).not.toContain('<a:lnTo>');
		const reloaded = shape?.customGeometryPaths ?? [];
		expect(reloaded).toHaveLength(2);
		expect(reloaded[0]?.segments.map((s) => s.type)).toEqual([
			'moveTo',
			'lineTo',
			'lineTo',
			'lineTo',
			'close',
		]);
		expect(reloaded[1]?.fillMode).toBe('none');
		expect(reloaded[1]?.stroke).toBe(false);
		expect(reloaded[1]?.segments.filter((s) => s.type === 'cubicBezTo')).toHaveLength(2);
	});

	it('writes a:custGeom in place of a:prstGeom, in schema order', async () => {
		const { custGeom, spPr } = await roundTrip({
			paths: [{ d: 'M0 0 L10 0 L5 10 Z', width: 10, height: 10 }],
		});
		expect(spPr).not.toContain('prstGeom');
		const order = ['<a:xfrm', '<a:custGeom', '<a:solidFill'].map((tag) => spPr.indexOf(tag));
		expect(order.every((index) => index >= 0)).toBe(true);
		expect([...order].sort((a, b) => a - b)).toEqual(order);
		const children = ['<a:avLst', '<a:gdLst', '<a:ahLst', '<a:cxnLst', '<a:rect', '<a:pathLst']
			.map((tag) => custGeom.indexOf(tag))
			.filter((index) => index >= 0);
		expect([...children].sort((a, b) => a - b)).toEqual(children);
		expect(custGeom).toContain('<a:pathLst>');
	});

	it('keeps the stroke and fill defaults when the flags are omitted', async () => {
		const { custGeom, shape } = await roundTrip({
			paths: [{ d: 'M0 0 L1 0 L1 1 Z', width: 1, height: 1 }],
		});
		expect(custGeom).not.toMatch(/<a:path [^>]*(?:fill|stroke)=/u);
		expect(shape?.customGeometryPaths?.[0]?.stroke ?? true).toBe(true);
		expect(shape?.customGeometryPaths?.[0]?.fillMode ?? 'norm').toBe('norm');
	});

	it('scales each path box over the shape, keeping fractional coordinates', () => {
		const el = createShapeElement('rect', {
			width: 200,
			height: 100,
			paths: [{ d: 'M0 0 L1 1 L0.5 0.25', width: 1, height: 1 }],
		});
		expect(el.shapeType).toBe('custom');
		const [path] = el.customGeometryPaths ?? [];
		expect(path?.segments[1]).toEqual({ type: 'lineTo', pt: { x: 1905000, y: 952500 } });
		expect(path?.segments[2]).toEqual({ type: 'lineTo', pt: { x: 952500, y: 238125 } });
	});

	it('takes segments instead of path data', () => {
		const el = createShapeElement('rect', {
			paths: [
				{
					segments: [
						{ type: 'moveTo', pt: { x: 0, y: 0 } },
						{ type: 'lineTo', pt: { x: 10, y: 0 } },
						{ type: 'close' },
					],
					width: 10,
					height: 10,
				},
			],
		});
		expect(el.customGeometryPaths?.[0]?.segments).toHaveLength(3);
		expect(el.pathData).toContain('M 0 0');
	});

	it('keeps the preset for an empty path list', () => {
		const el = createShapeElement('ellipse', { paths: [] });
		expect(el.shapeType).toBe('ellipse');
		expect(el.customGeometryPaths).toBeUndefined();
	});

	it('rejects a path it cannot write', () => {
		const make = (paths: ShapeOptions['paths']) => () => createShapeElement('rect', { paths });
		expect(make([{ width: 1, height: 1 }])).toThrow(RangeError);
		expect(make([{ d: 'M0 0 L1 1', width: 0, height: 1 }])).toThrow(RangeError);
		expect(make([{ d: 'M0 0 L1 1', width: Number.POSITIVE_INFINITY, height: 1 }])).toThrow(
			RangeError,
		);
		expect(
			make([
				{
					d: 'M0 0 L1 1',
					segments: [{ type: 'moveTo', pt: { x: 0, y: 0 } }],
					width: 1,
					height: 1,
				},
			]),
		).toThrow(RangeError);
		expect(
			make([{ segments: [{ type: 'moveTo', pt: { x: Number.NaN, y: 0 } }], width: 1, height: 1 }]),
		).toThrow(RangeError);
		expect(make([{ d: 'M0 0 L1', width: 1, height: 1 }])).toThrow(SyntaxError);
	});
});
