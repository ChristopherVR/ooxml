import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { attribute, children } from './sheet';
import { fixture, cell } from './test-fixtures';

it('uses distinct document drawing defaults for rectangle style inheritance', async () => {
	const bytes = await fixture({
		document: `<DocumentSettings DefaultLineStyle="1" DefaultFillStyle="2" DefaultTextStyle="3"/><StyleSheets><StyleSheet ID="1">${cell('LineColor', '#1860a8')}${cell('LineWeight', 0.05)}</StyleSheet><StyleSheet ID="2">${cell('FillForegnd', '#f0b050')}${cell('FillPattern', 1)}</StyleSheet><StyleSheet ID="3"/></StyleSheets>`,
	});
	const saved = await editVsdx(bytes, [
		{ type: 'create-rectangle', pageId: '0', shapeId: '2', x: 6.5, y: 2.5, width: 2, height: 1 },
	]);
	const model = await parseVsdx(saved.bytes);
	const created = model.pages[0]!.shapes.find((shape) => shape.id === '2')!;
	expect(created.style.fill).toBe('#f0b050');
	expect(created.style.lineColor).toBe('#1860a8');
	expect(created.style.lineWidth).toBe(0.05);
	const pkg = await VisioPackage.open(saved.bytes);
	const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
	const shape = children(children(root, 'Shapes')[0], 'Shape').find(
		(shape) => attribute(shape, 'ID') === '2',
	)!;
	expect(['LineStyle', 'FillStyle', 'TextStyle'].map((name) => attribute(shape, name))).toEqual([
		'1',
		'2',
		'3',
	]);
});

it('admits new rectangle defaults through the shared inherited-dependency guard', async () => {
	const bytes = await fixture({
		document: `<DocumentSettings DefaultFillStyle="2"/><StyleSheets><StyleSheet ID="2">${cell('FillForegnd', '#ffffff', 'Width')}</StyleSheet></StyleSheets>`,
	});
	await expect(
		editVsdx(bytes, [
			{ type: 'create-rectangle', pageId: '0', shapeId: '2', x: 6.5, y: 2.5, width: 2, height: 1 },
		]),
	).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY' });
});

for (const variable of [
	'VISIO_NATIVE_DRAW_DEFAULTS_DIR',
	'VISIO_NATIVE_DRAW_CUSTOM_DEFAULTS_DIR',
]) {
	const directory = process.env[variable];
	it.skipIf(!directory)(
		`matches native rectangle drawing defaults and poses (${variable})`,
		async () => {
			const bytes = await readFile(join(directory!, 'original.vsdx'));
			const native = await parseVsdx(bytes);
			const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
				rectangle: {
					shapeId: string;
					cells: Record<string, { value: number }>;
					transform: number[];
				};
			};
			const target = evidence.rectangle;
			const cleared = await editVsdx(bytes, [
				{ type: 'delete-shape', pageId: native.pages[0]!.id, shapeId: target.shapeId },
			]);
			const saved = await editVsdx(cleared.bytes, [
				{
					type: 'create-rectangle',
					pageId: native.pages[0]!.id,
					shapeId: target.shapeId,
					x: target.cells.PinX!.value,
					y: target.cells.PinY!.value,
					width: target.cells.Width!.value,
					height: target.cells.Height!.value,
				},
			]);
			const model = await parseVsdx(saved.bytes);
			const actual = model.pages[0]!.shapes.find((shape) => shape.id === target.shapeId)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === target.shapeId)!;
			expect(actual.style).toEqual(expected.style);
			expect(actual.geometry).toEqual(expected.geometry);
			for (let i = 0; i < 6; i++) expect(actual.transform[i]).toBeCloseTo(target.transform[i]!, 12);
			const before = await VisioPackage.open(bytes),
				after = await VisioPackage.open(saved.bytes);
			const originalRoot = await before.readXml('visio/pages/page1.xml', 'PageContents');
			const savedRoot = await after.readXml('visio/pages/page1.xml', 'PageContents');
			const originalShape = children(children(originalRoot, 'Shapes')[0], 'Shape').find(
				(shape) => attribute(shape, 'ID') === target.shapeId,
			)!;
			const savedShape = children(children(savedRoot, 'Shapes')[0], 'Shape').find(
				(shape) => attribute(shape, 'ID') === target.shapeId,
			)!;
			for (const category of ['LineStyle', 'FillStyle', 'TextStyle'])
				expect(attribute(savedShape, category)).toBe(attribute(originalShape, category));
			const matrixCells = (shape: Element) =>
				children(shape, 'Cell')
					.filter((node) => attribute(node, 'N')?.startsWith('QuickStyle'))
					.map((node) => [attribute(node, 'N'), attribute(node, 'V')]);
			expect(matrixCells(savedShape)).toEqual(matrixCells(originalShape));
			for (const path of before.paths())
				if (path !== 'visio/pages/page1.xml')
					expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		},
	);
}
