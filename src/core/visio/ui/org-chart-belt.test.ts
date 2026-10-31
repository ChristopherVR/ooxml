import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, parseVsdx, type VisioEdit, type VisioPage } from '../index';
import { visioOrgHierarchy } from './org-chart';
import { visioOrgAddPosition } from './org-chart-edit';
import { createTemplateVsdx } from './templates';

const pageOf = async (bytes: Uint8Array): Promise<VisioPage> => (await parseVsdx(bytes)).pages[0]!;
const apply = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
/** IDs of the template's positions: the group of each Belt master takes sixteen IDs. */
const T = (index: number) => String(1 + index * 16);
const centre = (page: VisioPage, id: string) => {
	const shape = page.shapes.find((item) => item.id === id)!;
	return {
		x: shape.transform[4] + shape.width / 2,
		y: shape.transform[5] + shape.height / 2,
		height: shape.height,
	};
};
/** The end points of the connector glued to `to` (its BeginX/BeginY/EndX/EndY caches). */
async function lineEnds(bytes: Uint8Array, to: string) {
	const page = await pageOf(bytes);
	const row = page.connectors.find((item) => item.toShapeId === to && item.fromCell === 'EndX')!;
	const xml = await (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
	const start = xml.indexOf(`<Shape ID="${row.fromShapeId}"`);
	const part = xml.slice(start, xml.indexOf('</Shape>', start));
	const cell = (name: string) => Number(new RegExp(`N="${name}" V="([^"]+)"`).exec(part)![1]);
	return {
		row,
		begin: { x: cell('BeginX'), y: cell('BeginY') },
		end: { x: cell('EndX'), y: cell('EndY') },
	};
}

describe('org chart Belt masters and reporting-line glue', () => {
	it('glues each reporting line to the manager Bottom and the subordinate Top point', async () => {
		const page = await pageOf(await createTemplateVsdx('organization-chart'));
		const rows = page.connectors.filter((row) => /^(Begin|End)X$/.test(row.fromCell));
		expect(rows).toHaveLength(10);
		for (const row of rows) {
			// The Belt masters' Connection rows are Top (0), Left, Right and Bottom (3).
			expect(row.toCell).toBe(/^Begin/.test(row.fromCell) ? 'Connections.X4' : 'Connections.X1');
		}
		expect(visioOrgHierarchy(page).childrenOf(T(1))).toEqual([T(4), T(5)]);
	});

	it('routes the line from the bottom of the manager to the top of the subordinate', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const page = await pageOf(bytes);
		const manager = centre(page, T(1));
		const leaf = centre(page, T(4));
		const { begin, end } = await lineEnds(bytes, T(4));
		expect(begin.x).toBeCloseTo(manager.x, 2);
		expect(begin.y).toBeCloseTo(manager.y - manager.height / 2, 2);
		expect(end.x).toBeCloseTo(leaf.x, 2);
		expect(end.y).toBeCloseTo(leaf.y + leaf.height / 2, 2);
	});

	it('keeps the line on the Top point when a position moves', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const movedBytes = await apply(bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: T(4), x: 9, y: 1.5 },
		]);
		const { row, end } = await lineEnds(movedBytes, T(4));
		expect(row.toCell).toBe('Connections.X1');
		const leaf = centre(await pageOf(movedBytes), T(4));
		expect(end.x).toBeCloseTo(leaf.x, 2);
		expect(end.y).toBeCloseTo(leaf.y + leaf.height / 2, 2);
	});

	it('turns a Position Belt that gains its first subordinate into a Manager Belt', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const plan = visioOrgAddPosition(await pageOf(bytes), T(4), 'subordinate')!;
		const after = await apply(bytes, plan.edits);
		const page = await pageOf(after);
		const hierarchy = visioOrgHierarchy(page);
		expect(page.shapes.some((shape) => shape.id === T(4))).toBe(false);
		const promoted = hierarchy.parentOf(plan.newId)!;
		const shape = page.shapes.find((item) => item.id === promoted)!;
		expect(shape.name).toMatch(/^Manager Belt/);
		expect(shape.text.plainText).toBe('Name\nPosition');
		expect(shape.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Position');
		// Still under the same manager, with its line glued to the new shape's Top point.
		expect(hierarchy.parentOf(promoted)).toBe(T(1));
		expect(hierarchy.positions).toHaveLength(7);
		for (const row of page.connectors)
			expect(page.shapes.some((item) => item.id === row.toShapeId)).toBe(true);
		const incoming = await lineEnds(after, promoted);
		expect(incoming.row.toCell).toBe('Connections.X1');
		const at = centre(page, promoted);
		expect(incoming.end.y).toBeCloseTo(at.y + at.height / 2, 2);
		// The added subordinate stays a Position Belt.
		expect(page.shapes.find((item) => item.id === plan.newId)?.name).toMatch(/^Position Belt/);
	});

	it('carries every Shape Data row of the position to the Manager Belt', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const custom = await apply(bytes, [
			{
				type: 'set-shape-data',
				pageId: '0',
				shapeId: T(4),
				row: 'Cost_Center',
				data: {
					label: 'Cost Center',
					prompt: 'Which one',
					type: 'number',
					format: '0.0',
					value: '42',
					invisible: true,
				},
			},
			{
				type: 'set-shape-data',
				pageId: '0',
				shapeId: T(4),
				row: 'Started',
				data: {
					label: 'Started',
					prompt: '',
					type: 'date',
					format: '',
					value: '2024-03-05',
				},
			},
			{
				type: 'set-shape-data',
				pageId: '0',
				shapeId: T(4),
				row: 'Email',
				data: { label: 'E-Mail', prompt: '', type: 'string', format: '', value: 'a@b.test' },
			},
		]);
		const plan = visioOrgAddPosition(await pageOf(custom), T(4), 'subordinate')!;
		const page = await pageOf(await apply(custom, plan.edits));
		const promoted = page.shapes.find(
			(item) => item.id === visioOrgHierarchy(page).parentOf(plan.newId),
		)!;
		const row = (name: string) => promoted.shapeData?.find((item) => item.name === name);
		expect(row('Cost_Center')).toMatchObject({
			label: 'Cost Center',
			prompt: 'Which one',
			valueKind: 'number',
			value: 42,
			invisible: true,
		});
		expect(row('Started')?.valueKind).toBe('date');
		expect(row('Started')?.rawValue).toBe('2024-03-05T00:00:00');
		expect(row('Email')?.rawValue).toBe('a@b.test');
	});

	it('keeps the glued point of a line recorded with Visio named Connection rows', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const zip = await JSZip.loadAsync(bytes);
		const path = 'visio/pages/page1.xml';
		const named = (await zip.file(path)!.async('string'))
			.replaceAll('ToCell="Connections.X4"', 'ToCell="Connections.Bottom.X"')
			.replaceAll('ToCell="Connections.X1"', 'ToCell="Connections.Top.X"');
		zip.file(path, named);
		const edited = await zip.generateAsync({ type: 'uint8array' });
		const before = await pageOf(edited);
		const plan = visioOrgAddPosition(before, T(4), 'subordinate')!;
		const drop = plan.edits.find((edit) => edit.type === 'drop-stencil-master')!;
		const onManager = plan.edits.filter(
			(edit) => edit.type === 'glue-connector' && edit.target === drop.shapeId,
		);
		// The line that entered the leaf's Top point enters the Manager Belt's Top point (row 0),
		// and the new line leaves its Bottom point (row 3).
		expect(
			onManager.map((edit) => (edit.type === 'glue-connector' ? edit.point : -1)).sort(),
		).toEqual([0, 3]);
	});

	it('leaves a Manager Belt a Manager Belt and an added coworker a Position Belt', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const before = await pageOf(bytes);
		const sub = visioOrgAddPosition(before, T(1), 'subordinate')!;
		const afterSub = await pageOf(await apply(bytes, sub.edits));
		expect(afterSub.shapes.find((shape) => shape.id === T(1))?.name).toMatch(/^Manager Belt/);
		const peer = visioOrgAddPosition(before, T(1), 'peer')!;
		const afterPeer = await pageOf(await apply(bytes, peer.edits));
		expect(afterPeer.shapes.find((shape) => shape.id === peer.newId)?.name).toMatch(
			/^Position Belt/,
		);
	});

	it('does not turn a Position Belt into a Manager Belt for an assistant', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const plan = visioOrgAddPosition(await pageOf(bytes), T(4), 'assistant')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		expect(page.shapes.find((shape) => shape.id === T(4))?.name).toMatch(/^Position Belt/);
	});

	it('resizes a Belt master instance, keeping the saved text extent for the text pass', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const result = await editVsdx(bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: T(4), width: 2, height: 1.2 },
		]);
		const page = await pageOf(result.bytes);
		const shape = page.shapes.find((item) => item.id === T(4))!;
		expect([shape.width, shape.height]).toEqual([2, 1.2]);
		// The belt and the photo frame follow the new size, as Visio recalculates them.
		const widths: number[] = [];
		const walk = (list: typeof page.shapes) =>
			list.forEach((item) => {
				widths.push(item.width);
				walk(item.children);
			});
		walk(shape.children);
		expect(widths.some((width) => Math.abs(width - 2) < 1e-6)).toBe(true);
		// The photo frame's PinY is the value Visio 16 computes for a 2 x 1.2 in Position Belt
		// (read over COM: MAX(Height*0.5, BeltSize*1.5 + frame/2 + LineWeight/2) = 0.6338172572).
		const xml = await (
			await JSZip.loadAsync(result.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		const frameXml = xml.slice(xml.indexOf('<Shape ID="74"'));
		expect(Number(/N="PinY" V="([^"]+)"/.exec(frameXml)![1])).toBeCloseTo(0.6338172572, 9);
		expect(result.diagnostics.map((item) => item.code)).toContain('edit-text-size-kept');
		// The glued reporting line still ends on the resized position's Top point.
		const line = await lineEnds(result.bytes, T(4));
		const at = centre(page, T(4));
		expect(line.end.y).toBeCloseTo(at.y + at.height / 2, 2);
	});
});
