import { describe, expect, it } from 'vitest';
import { editVsdx, parseVsdx, type VisioEdit, type VisioPage } from '../index';
import {
	VISIO_ORG_ROWS,
	visioIsOrgPosition,
	visioOrgAddPosition,
	visioOrgDeletePosition,
	visioOrgHierarchy,
	visioOrgReLayout,
} from './org-chart';
import { createTemplateVsdx } from './templates';

const pageOf = async (bytes: Uint8Array): Promise<VisioPage> => (await parseVsdx(bytes)).pages[0]!;
const apply = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
const centre = (page: VisioPage, id: string) => {
	const shape = page.shapes.find((item) => item.id === id)!;
	return { x: shape.transform[4] + shape.width / 2, y: shape.transform[5] + shape.height / 2 };
};

describe('org chart', () => {
	it('writes Visio org chart Shape Data on every position of the template', async () => {
		const page = await pageOf(await createTemplateVsdx('organization-chart'));
		const hierarchy = visioOrgHierarchy(page);
		expect(hierarchy.positions.map((shape) => shape.id)).toEqual(['1', '2', '3', '4', '5', '6']);
		const top = page.shapes.find((shape) => shape.id === '1')!;
		expect(top.shapeData?.map((item) => item.name)).toEqual(VISIO_ORG_ROWS.map((r) => r.row));
		expect(top.shapeData?.find((item) => item.name === 'Email')?.label).toBe('E-Mail');
		expect(top.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Chief Executive');
		expect(hierarchy.childrenOf('1')).toEqual(['2', '3', '4']);
		expect(hierarchy.parentOf('5')).toBe('2');
		expect(visioIsOrgPosition(page.shapes.find((shape) => shape.id === '100'))).toBe(false);
	});

	it('adds a subordinate with data and a glued connector in one batch', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const plan = visioOrgAddPosition(await pageOf(bytes), '3', 'subordinate')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		const hierarchy = visioOrgHierarchy(page);
		expect(hierarchy.parentOf(plan.newId)).toBe('3');
		const added = page.shapes.find((shape) => shape.id === plan.newId)!;
		expect(added.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Position');
		// The tree was laid out again: the new box sits below its manager and overlaps nothing.
		expect(centre(page, plan.newId).y).toBeLessThan(centre(page, '3').y);
		const ids = hierarchy.positions.map((shape) => shape.id);
		for (const a of ids)
			for (const b of ids)
				if (a < b) {
					const [p, q] = [centre(page, a), centre(page, b)];
					expect(Math.abs(p.x - q.x) > 1.7 || Math.abs(p.y - q.y) > 0.7).toBe(true);
				}
	});

	it('adds a coworker under the same manager and refuses one for the top position', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const before = await pageOf(bytes);
		expect(visioOrgAddPosition(before, '1', 'peer')).toBeUndefined();
		const plan = visioOrgAddPosition(before, '5', 'peer')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		expect(visioOrgHierarchy(page).parentOf(plan.newId)).toBe('2');
	});

	it('deletes a position with its connectors and reattaches its subordinates', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const batches = visioOrgDeletePosition(await pageOf(bytes), '2')!;
		let current = bytes;
		for (const batch of batches) current = await apply(current, batch);
		const page = await pageOf(current);
		const hierarchy = visioOrgHierarchy(page);
		expect(page.shapes.some((shape) => shape.id === '2')).toBe(false);
		expect(hierarchy.positions).toHaveLength(5);
		expect(hierarchy.parentOf('5')).toBe('1');
		expect(hierarchy.parentOf('6')).toBe('1');
		expect(page.connectors.every((c) => page.shapes.some((s) => s.id === c.toShapeId))).toBe(true);
	});

	it('lays a scattered tree out again', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const moved = await apply(bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: '5', x: 9, y: 1 },
		]);
		const plan = visioOrgReLayout(await pageOf(moved));
		expect(plan?.length).toBeGreaterThan(0);
		const page = await pageOf(await apply(moved, plan!));
		expect(centre(page, '5').y).toBeLessThan(centre(page, '2').y);
		expect(Math.abs(centre(page, '5').x - 9)).toBeGreaterThan(1);
	});
});
