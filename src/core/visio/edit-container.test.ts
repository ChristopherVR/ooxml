import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture } from './test-fixtures';
import { attribute, children } from './sheet';
import { containerMemberIds } from './shape-structure';
import { VISIO_CONTAINER_HEADING, VISIO_CONTAINER_MARGIN } from './edit-diagram-parts-commands';

const rect = (shapeId: string, x: number, y: number, width = 1, height = 1): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width,
	height,
});
const insert = (memberIds: string[]): VisioEdit =>
	({
		type: 'insert-container',
		pageId: '0',
		shapeId: '9',
		memberIds,
		style: 'plain',
		heading: 'Team',
	}) as VisioEdit;
async function base(): Promise<Uint8Array> {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	return (await editVsdx(blank, [rect('1', 2, 2), rect('2', 4, 2), rect('3', 10, 8)])).bytes;
}
async function framed(): Promise<Uint8Array> {
	return (await editVsdx(await base(), [insert(['1', '2'])])).bytes;
}
async function pageRoot(bytes: Uint8Array): Promise<Element> {
	return (await VisioPackage.open(bytes)).readXml('visio/pages/page1.xml');
}
const sheet = (root: Element, id: string) =>
	Array.from(root.getElementsByTagName('Shape')).find((node) => attribute(node, 'ID') === id)!;
const cell = (root: Element, id: string, name: string) =>
	children(sheet(root, id), 'Cell').find((node) => attribute(node, 'N') === name);
const formula = (root: Element, id: string) => attribute(cell(root, id, 'Relationships'), 'F');
const value = (root: Element, id: string, name: string) =>
	Number(attribute(cell(root, id, name), 'V'));
const members = async (bytes: Uint8Array) => {
	const shape = (await parseVsdx(bytes)).pages[0]!.shapes.find((item) => item.id === '9');
	return shape?.structure?.type === 'container' ? shape.structure.memberIds : undefined;
};
const reconcile = (...shapeIds: string[]): VisioEdit => ({
	type: 'reconcile-containers',
	pageId: '0',
	shapeIds,
});
const move = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId,
	x,
	y,
});

describe('container membership in Visio form', () => {
	it('writes DEPENDSON relationships, back references and msvSDContainer rows', async () => {
		const root = await pageRoot(await framed());
		expect(formula(root, '9')).toBe('SUM(DEPENDSON(4,Sheet.1!SheetRef(),Sheet.2!SheetRef()))');
		expect(formula(root, '1')).toBe('SUM(DEPENDSON(1,Sheet.9!SheetRef()))');
		expect(formula(root, '2')).toBe('SUM(DEPENDSON(1,Sheet.9!SheetRef()))');
		expect(formula(root, '3')).toBeUndefined();
		const rows = Array.from(sheet(root, '9').getElementsByTagName('Row')).map((row) =>
			attribute(row, 'N'),
		);
		expect(rows).toEqual(
			expect.arrayContaining(['msvStructureType', 'msvSDContainerMargin', 'msvSDContainerResize']),
		);
		expect(rows).not.toContain('ooxmlContainerMembers');
		expect(await members(await framed())).toEqual(['1', '2']);
	});

	it('still reads membership from the private row of older files', () => {
		expect(containerMemberIds('1,2', undefined)).toEqual(['1', '2']);
		expect(containerMemberIds(undefined, 'SUM(DEPENDSON(4,Sheet.3!SheetRef()))')).toEqual(['3']);
	});

	it('moves a container with the relationships in place and round-trips', async () => {
		const saved = await editVsdx(await framed(), [move('9', 5, 5)]);
		expect(await members(saved.bytes)).toEqual(['1', '2']);
	});

	it('does not treat the Relationships links as dynamic dependencies', async () => {
		const saved = await editVsdx(await framed(), [
			rect('5', 8, 8),
			move('1', 2.2, 2),
			reconcile('1', '5'),
		]);
		expect(await members(saved.bytes)).toEqual(['1', '2']);
	});

	it('replaces members and keeps both sides in step', async () => {
		const saved = await editVsdx(await framed(), [
			{ type: 'set-container-members', pageId: '0', shapeId: '9', memberIds: ['2', '3'] },
		]);
		const root = await pageRoot(saved.bytes);
		expect(formula(root, '9')).toBe('SUM(DEPENDSON(4,Sheet.2!SheetRef(),Sheet.3!SheetRef()))');
		expect(formula(root, '1')).toBeUndefined();
		expect(formula(root, '3')).toBe('SUM(DEPENDSON(1,Sheet.9!SheetRef()))');
		expect(await members(saved.bytes)).toEqual(['2', '3']);
	});

	it('refuses members that are not top-level shapes of the page', async () => {
		await expect(
			editVsdx(await framed(), [
				{ type: 'set-container-members', pageId: '0', shapeId: '9', memberIds: ['77'] },
			]),
		).rejects.toThrow();
	});

	it('removes a deleted member from the relationships', async () => {
		const saved = await editVsdx(await framed(), [
			{ type: 'delete-shape', pageId: '0', shapeId: '2' },
		]);
		expect(await members(saved.bytes)).toEqual(['1']);
		expect(formula(await pageRoot(saved.bytes), '9')).toBe('SUM(DEPENDSON(4,Sheet.1!SheetRef()))');
	});
});

describe('reconcile-containers', () => {
	it('adds a shape dropped fully inside and removes one moved out', async () => {
		// Shape 3 moves into the frame (members span x 1.5..4.5, y 1.5..2.5).
		const into = await editVsdx(await framed(), [move('3', 3, 2), reconcile('3')]);
		expect(await members(into.bytes)).toEqual(['1', '2', '3']);
		expect(formula(await pageRoot(into.bytes), '3')).toBe('SUM(DEPENDSON(1,Sheet.9!SheetRef()))');
		const out = await editVsdx(into.bytes, [move('3', 30, 30), reconcile('3')]);
		expect(await members(out.bytes)).toEqual(['1', '2']);
		expect(formula(await pageRoot(out.bytes), '3')).toBeUndefined();
	});

	it('grows the container to fit a member dragged to the edge, with its margin', async () => {
		const before = await pageRoot(await framed());
		const saved = await editVsdx(await framed(), [move('2', 4.5, 2), reconcile('2')]);
		const root = await pageRoot(saved.bytes);
		const right = value(root, '9', 'PinX') + value(root, '9', 'Width') / 2;
		expect(right).toBeCloseTo(5 + VISIO_CONTAINER_MARGIN, 9);
		expect(value(root, '9', 'Height')).toBeCloseTo(value(before, '9', 'Height'), 9);
		expect(value(root, '9', 'PinX') - value(root, '9', 'Width') / 2).toBeCloseTo(
			value(before, '9', 'PinX') - value(before, '9', 'Width') / 2,
			9,
		);
		expect(await members(saved.bytes)).toEqual(['1', '2']);
	});

	it('grows upwards over the heading allowance', async () => {
		const saved = await editVsdx(await framed(), [move('1', 2, 3), reconcile('1')]);
		const root = await pageRoot(saved.bytes);
		const top = value(root, '9', 'PinY') + value(root, '9', 'Height') / 2;
		expect(top).toBeCloseTo(3.5 + VISIO_CONTAINER_MARGIN + VISIO_CONTAINER_HEADING, 9);
	});

	it('leaves membership alone when the container moved with its members', async () => {
		const saved = await editVsdx(await framed(), [
			move('9', 6, 6),
			move('1', 5, 6),
			move('2', 7, 6),
			reconcile('9', '1', '2'),
		]);
		expect(await members(saved.bytes)).toEqual(['1', '2']);
	});
});
