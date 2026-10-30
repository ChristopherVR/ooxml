import { describe, expect, it } from 'vitest';
import { parseVsdx } from '../index';
import { VISIO_TEMPLATES, createTemplateVsdx, isVisioTemplateId } from './templates';

/**
 * Visio's own org chart masters drive colours, line patterns and rounding from the theme and
 * cache the symbolic result, which the viewer reports as a warning instead of resolving.
 */
const THEME_DRIVEN = new Set([
	'missing-cached-value',
	'unsupported-color',
	'unresolved-line-pattern',
	'unresolved-line-cap',
	'unsupported-image-clipping',
	'master-resize-cached-values',
	'unsupported-theme-shadow',
]);

describe('drawing templates', () => {
	it('lists Blank Drawing first', () => {
		expect(VISIO_TEMPLATES[0]).toMatchObject({ id: 'blank', name: 'Blank Drawing' });
		expect(VISIO_TEMPLATES.map((template) => template.name)).toEqual([
			'Blank Drawing',
			'Basic Flowchart',
			'Cross-Functional Flowchart',
			'Cross-Functional Flowchart (Vertical)',
			'Basic Organization Chart',
		]);
		expect(isVisioTemplateId('nope')).toBe(false);
	});

	it('builds a blank drawing', async () => {
		const model = await parseVsdx(await createTemplateVsdx('blank'));
		expect(model.pages[0]!.shapes).toEqual([]);
	});

	it.each([
		['basic-flowchart', 5, 4],
		['cross-functional-flowchart', 8, 3],
		['cross-functional-flowchart-vertical', 8, 3],
		['organization-chart', 6, 5],
	] as const)('round-trips %s with glued connectors', async (id, boxes, links) => {
		const model = await parseVsdx(await createTemplateVsdx(id));
		const page = model.pages[0]!;
		const warnings = model.diagnostics.filter(
			(d) => d.severity === 'warning' && !(id === 'organization-chart' && THEME_DRIVEN.has(d.code)),
		);
		expect(warnings).toEqual([]);
		// Each link is a connector shape with a begin and an end glue.
		expect(page.connectors).toHaveLength(links * 2);
		const ids = new Set(page.shapes.map((shape) => shape.id));
		for (const connect of page.connectors) {
			expect(ids.has(connect.fromShapeId)).toBe(true);
			expect(ids.has(connect.toShapeId)).toBe(true);
		}
		expect(page.shapes.length).toBeGreaterThanOrEqual(boxes + links);
	});

	it('builds real swimlanes whose shapes belong to their lane', async () => {
		const page = (await parseVsdx(await createTemplateVsdx('cross-functional-flowchart')))
			.pages[0]!;
		const names = page.shapes.map((shape) => shape.text?.plainText);
		expect(names).toEqual(expect.arrayContaining(['Customer', 'Sales', 'Fulfilment']));
		const structure = page.shapes.find((shape) => shape.id === '50')?.structure;
		expect(structure).toMatchObject({ swimlane: { role: 'frame' } });
		const members = (id: string) => {
			const found = page.shapes.find((shape) => shape.id === id)?.structure;
			return found?.type === 'container' ? found.memberIds : [];
		};
		// The frame holds Visio's Phase List and Swimlane List; the lanes are the latter's members.
		const lists = members('50').map(
			(id) => page.shapes.find((shape) => shape.id === id)?.structure,
		);
		expect(lists.map((list) => list?.type)).toEqual(['list', 'list']);
		expect(
			lists.flatMap((list) =>
				list?.type === 'list' && list.kind === 'swimlane' ? list.memberIds : [],
			),
		).toEqual(['10', '11', '12']);
		expect(members('10')).toContain('1');
		expect(members('11')).toEqual(expect.arrayContaining(['2', '3']));
		expect(members('12')).toContain('4');
	});

	it('builds a vertical flowchart: columns side by side, shapes in their column', async () => {
		const page = (await parseVsdx(await createTemplateVsdx('cross-functional-flowchart-vertical')))
			.pages[0]!;
		const frame = page.shapes.find((shape) => shape.id === '50')?.structure;
		expect(frame).toMatchObject({
			swimlane: { role: 'frame', options: { direction: 'vertical' } },
		});
		const lane = (id: string) => page.shapes.find((shape) => shape.id === id)!;
		expect(lane('10').width).toBeLessThan(lane('10').height);
		const members = (id: string) => {
			const found = lane(id).structure;
			return found?.type === 'container' ? found.memberIds : [];
		};
		expect(members('10')).toContain('1');
		expect(members('11')).toEqual(expect.arrayContaining(['2', '3']));
		expect(members('12')).toContain('4');
		// Left to right in the order the lanes are listed.
		expect(lane('10').transform[4]).toBeLessThan(lane('11').transform[4]);
		expect(lane('11').transform[4]).toBeLessThan(lane('12').transform[4]);
	});

	it('rejects unknown ids', async () => {
		await expect(createTemplateVsdx('x' as never)).rejects.toThrow('Unknown');
	});
});
