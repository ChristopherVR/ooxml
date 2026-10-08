import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { analyzeVisioFormula, evaluateVisioFormula, type VisioFormulaValue } from './formula';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

const directory = process.env.VISIO_NATIVE_ANCHORED_RESIZE_DIR;
interface NativeCell {
	formula: string;
	value: number;
}
interface NativeCase {
	pageId: string;
	shapeId: string;
	mode: string;
	direction: number;
	scale: number;
	error: string | null;
	before: Record<string, NativeCell>;
	after: Record<string, NativeCell>;
	beforeTransform: number[];
	afterTransform: number[];
}
const anchors = [
	[0, 0.5],
	[0, 0],
	[0.5, 0],
	[1, 0],
	[1, 0.5],
	[1, 1],
	[0.5, 1],
	[0, 1],
] as const;
const supported = new Set([
	'half',
	'quarter',
	'constant',
	'guard-locpin',
	'offset-locpin',
	'rotated',
	'flip-x',
	'flip-y',
	'guard-flip',
	'pin-formula',
	'pin-dependent-locpin',
	'lock-aspect',
]);
const value = (cells: Record<string, NativeCell>, name: string) => cells[name]!.value;
function localPin(item: NativeCase, axis: 'X' | 'Y'): number {
	const dimension = axis === 'X' ? 'Width' : 'Height';
	const source = item.before[`LocPin${axis}`]!;
	if (!analyzeVisioFormula(source.formula).guarded)
		return (source.value / value(item.before, dimension)) * value(item.after, dimension);
	const result = evaluateVisioFormula(source.formula, (reference): VisioFormulaValue => {
		if (reference.shapeId || !['Width', 'Height'].includes(reference.cell))
			throw new Error('Unproven guarded local pin');
		return { value: value(item.after, reference.cell), unit: 'length' };
	});
	return result.value;
}
/** Read-only source math prototype. No production anchored edit is admitted by these tests. */
describe.skipIf(!directory)('native anchored resize prototype', () => {
	it('derives pins from normalized unguarded or evaluated guarded local pins', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as { cases: NativeCase[] };
		let measured = 0;
		for (const item of evidence.cases) {
			if (!supported.has(item.mode) || item.error) continue;
			const [anchorX, anchorY] = anchors[item.direction]!;
			const [a, b, c, d] = item.beforeTransform;
			const localX = localPin(item, 'X'),
				localY = localPin(item, 'Y');
			expect(localX, item.mode).toBeCloseTo(value(item.after, 'LocPinX'), 10);
			expect(localY, item.mode).toBeCloseTo(value(item.after, 'LocPinY'), 10);
			const deltaX =
				anchorX * (value(item.before, 'Width') - value(item.after, 'Width')) -
				value(item.before, 'LocPinX') +
				localX;
			const deltaY =
				anchorY * (value(item.before, 'Height') - value(item.after, 'Height')) -
				value(item.before, 'LocPinY') +
				localY;
			expect(
				value(item.before, 'PinX') + a! * deltaX + c! * deltaY,
				`${item.mode}/${item.direction}/${item.scale}/PinX`,
			).toBeCloseTo(value(item.after, 'PinX'), 10);
			expect(
				value(item.before, 'PinY') + b! * deltaX + d! * deltaY,
				`${item.mode}/${item.direction}/${item.scale}/PinY`,
			).toBeCloseTo(value(item.after, 'PinY'), 10);
			++measured;
		}
		expect(measured).toBeGreaterThan(0);
	});
	it('records protection and post-dimension transform changes as distinct unsupported cases', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as { cases: NativeCase[] };
		for (const item of evidence.cases) {
			if (
				item.mode === 'lock-move' ||
				(item.mode === 'lock-width' && ![2, 6].includes(item.direction))
			) {
				expect(item.error, item.mode).toBeTruthy();
				expect(item.after.Width).toEqual(item.before.Width);
				expect(item.after.PinX).toEqual(item.before.PinX);
			}
			if (item.mode === 'guard-pin') expect(item.after.PinX).toEqual(item.before.PinX);
			if (item.mode === 'guard-constant-locpin' && item.direction === 0) {
				// Native pin movement uses the old local-pin ratio even when GUARD retains a constant.
				expect(item.after.LocPinX).toEqual(item.before.LocPinX);
				expect(value(item.after, 'PinX')).not.toBe(value(item.before, 'PinX'));
			}
		}
	});
});

describe.skipIf(!directory)('native anchored source-edit oracle', () => {
	it('matches supported native scenes and cache values in the actual saved package', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as { cases: NativeCase[] };
		const cases = evidence.cases.filter((item) => supported.has(item.mode) && !item.error);
		const bytes = new Uint8Array(await readFile(join(directory!, 'original.vsdx')));
		const saved = await editVsdx(
			bytes,
			cases.map((item) => {
				const [x, y] = anchors[item.direction]!;
				return {
					type: 'resize-shape' as const,
					pageId: item.pageId,
					shapeId: item.shapeId,
					width: value(item.after, 'Width'),
					height: value(item.after, 'Height'),
					anchor: { x, y },
				};
			}),
			{ limits: { maxRuntimeMs: 120_000 } },
		);
		const actual = await parseVsdx(saved.bytes),
			native = await parseVsdx(new Uint8Array(await readFile(join(directory!, 'native.vsdx'))));
		for (const item of cases) {
			const shape = actual.pages.find((page) => page.id === item.pageId)!.shapes[0]!;
			const reference = native.pages.find((page) => page.id === item.pageId)!.shapes[0]!;
			expect(shape.geometry, `${item.mode}/${item.direction}`).toEqual(reference.geometry);
			expect(shape.style).toEqual(reference.style);
			expect(shape.text.plainText).toEqual(reference.text.plainText);
			expect(shape.text.width).toBeCloseTo(reference.text.width, 10);
			expect(shape.text.height).toBeCloseTo(reference.text.height, 10);
			shape.text.transform.forEach((value, index) =>
				expect(value).toBeCloseTo(reference.text.transform[index]!, 10),
			);
			shape.transform.forEach((value, index) =>
				expect(value).toBeCloseTo(reference.transform[index]!, 10),
			);
		}
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		for (const [path, entry] of Object.entries(before.files))
			if (!entry.dir && !saved.changedParts.includes(path))
				expect(await after.file(path)!.async('uint8array'), path).toEqual(
					await entry.async('uint8array'),
				);
		await writeFile(join(directory!, 'core.vsdx'), saved.bytes);
		await writeFile(join(directory!, 'accepted-cases.json'), JSON.stringify(cases));
	}, 120_000);
});
