import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { evaluateVisioFormula, type VisioFormulaValue } from './formula';
import { VisioPackage } from './package';
import { attribute, children } from './sheet';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

// Visio 16 DrawLine caches captured by scripts/record-visio-line-movement.ps1.
it.each([
	[3, 1.5, 0],
	[2.732050807568877, 2.5, Math.PI / 6],
	[1, 3.5, Math.PI / 2],
	[-1, 1.5, Math.PI],
])('matches native line length, midpoint and angle for end (%s, %s)', (x, y, angle) => {
	for (const [dx, dy] of [
		[0, 0],
		[2.25, 1.5],
	] as const) {
		const values: Record<string, number> = {
			BeginX: 1 + dx,
			BeginY: 1.5 + dy,
			EndX: x! + dx,
			EndY: y! + dy,
		};
		const resolve = ({ cell }: { cell: string }): VisioFormulaValue => {
			if (!Object.hasOwn(values, cell)) throw new Error(`Missing ${cell}`);
			return { value: values[cell]!, unit: 'length' };
		};
		expect(
			evaluateVisioFormula('SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)', resolve).value,
		).toBeCloseTo(2, 12);
		const rotation = evaluateVisioFormula('ATAN2(EndY-BeginY,EndX-BeginX)', resolve);
		expect(rotation.unit).toBe('angle');
		expect(rotation.value).toBeCloseTo(angle!, 12);
		expect(evaluateVisioFormula('(BeginX+EndX)/2', resolve).value).toBeCloseTo(
			(1 + x!) / 2 + dx,
			12,
		);
		expect(evaluateVisioFormula('(BeginY+EndY)/2', resolve).value).toBeCloseTo(
			(1.5 + y!) / 2 + dy,
			12,
		);
	}
});

const directory = process.env.VISIO_NATIVE_LINE_MOVEMENT_DIR;
const resizeDirectory = process.env.VISIO_NATIVE_LINE_RESIZE_DIR;
it.skipIf(!resizeDirectory)(
	'matches native Width-cell resize caches and XYToPage poses',
	async () => {
		const bytes = await readFile(join(resizeDirectory!, 'moved.vsdx'));
		const original = await parseVsdx(bytes);
		const evidence = JSON.parse(
			await readFile(join(resizeDirectory!, 'evidence.json'), 'utf8'),
		) as {
			cases: {
				shapeId: string;
				resized: Record<string, { value: number; formula: string }>;
				resizedTransform: number[];
			}[];
		};
		const saved = await editVsdx(
			bytes,
			evidence.cases.map((item) => ({
				type: 'resize-shape',
				pageId: original.pages[0]!.id,
				shapeId: item.shapeId,
				width: item.resized.Width!.value,
				height: item.resized.Height!.value,
			})),
		);
		const result = await parseVsdx(saved.bytes);
		const pkg = await VisioPackage.open(saved.bytes);
		const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
		for (const item of evidence.cases) {
			const node = children(children(root, 'Shapes')[0], 'Shape').find(
				(node) => attribute(node, 'ID') === item.shapeId,
			);
			const cells = new Map(children(node, 'Cell').map((cell) => [attribute(cell, 'N'), cell]));
			for (const [name, native] of Object.entries(item.resized)) {
				expect(Number(attribute(cells.get(name), 'V')), `${item.shapeId}/${name}`).toBeCloseTo(
					native.value,
					12,
				);
				const formula = attribute(cells.get(name), 'F');
				if (formula) expect(formula, name).toBe(native.formula);
			}
			expect(attribute(cells.get('Width'), 'F')).toBeUndefined();
			const shape = result.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
			for (let i = 0; i < 6; i++)
				expect(shape.transform[i]).toBeCloseTo(item.resizedTransform[i]!, 12);
		}
		const native = await parseVsdx(await readFile(join(resizeDirectory!, 'resized.vsdx')));
		expect(result.pages[0]!.shapes).toEqual(native.pages[0]!.shapes);
	},
);
it.skipIf(!directory)(
	'moves genuine native lines with the same caches as native endpoint translation',
	async () => {
		const original = await readFile(join(directory!, 'original.vsdx'));
		const parsed = await parseVsdx(original);
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: { shapeId: string; after: Record<string, { value: number; formula: string }> }[];
		};
		const saved = await editVsdx(
			original,
			evidence.cases.map((item) => ({
				type: 'move-shape',
				pageId: parsed.pages[0]!.id,
				shapeId: item.shapeId,
				x: item.after.PinX!.value,
				y: item.after.PinY!.value,
			})),
		);
		const pkg = await VisioPackage.open(saved.bytes);
		const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
		for (const item of evidence.cases) {
			const shape = children(children(root, 'Shapes')[0], 'Shape').find(
				(node) => attribute(node, 'ID') === item.shapeId,
			);
			const cells = new Map(children(shape, 'Cell').map((node) => [attribute(node, 'N'), node]));
			for (const [name, native] of Object.entries(item.after)) {
				expect(Number(attribute(cells.get(name), 'V')), `${item.shapeId}/${name}`).toBeCloseTo(
					native.value,
					12,
				);
				const formula = attribute(cells.get(name), 'F');
				if (formula) expect(formula, name).toBe(native.formula);
			}
		}
		const result = await parseVsdx(saved.bytes);
		expect(result.pages[0]!.shapes).toHaveLength(4);
		for (const shape of result.pages[0]!.shapes) expect(shape.width).toBeCloseTo(2, 12);
	},
);
it.skipIf(!directory)(
	'matches native saved original and moved line caches and preserves transform formulas',
	async () => {
		type Cell = { value: number; formula: string };
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			application: string;
			cases: {
				shapeId: string;
				oneD: number;
				before: Record<string, Cell>;
				after: Record<string, Cell>;
				delta: number[];
			}[];
		};
		expect(evidence.application).toBe('Microsoft Visio');
		expect(evidence.cases).toHaveLength(4);
		for (const state of ['before', 'after'] as const) {
			const pkg = await VisioPackage.open(
				await readFile(join(directory!, state === 'before' ? 'original.vsdx' : 'moved.vsdx')),
			);
			const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
			const shapes = children(children(root, 'Shapes')[0], 'Shape');
			expect(shapes).toHaveLength(4);
			for (const item of evidence.cases) {
				expect(item.oneD).not.toBe(0);
				const shape = shapes.find((node) => attribute(node, 'ID') === item.shapeId);
				expect(shape).toBeDefined();
				const cells = new Map(children(shape, 'Cell').map((node) => [attribute(node, 'N'), node]));
				for (const [name, native] of Object.entries(item[state])) {
					const node = cells.get(name);
					expect(node, name).toBeDefined();
					expect(Number(attribute(node, 'V')), name).toBeCloseTo(native.value, 12);
					const formula = attribute(node, 'F');
					if (formula) expect(formula, name).toBe(native.formula);
				}
				for (const name of ['PinX', 'PinY', 'Width', 'Angle', 'LocPinX', 'LocPinY']) {
					expect(attribute(cells.get(name), 'F'), name).toBe(item.before[name]!.formula);
					expect(item.after[name]!.formula, name).toBe(item.before[name]!.formula);
				}
			}
		}
	},
);
