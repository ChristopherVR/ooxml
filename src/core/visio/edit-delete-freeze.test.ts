import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { cell, fixture, rectangle, shape } from './test-fixtures';

// Visio replaces the formula of a cell that read a deleted shape by its last result as a constant
// (evidence/native-delete-error-2026-10-29.json: `Sheet.2!Width*0.25` becomes `0.5 in`).
const leaf = (id: string, extra = '') =>
	shape(
		id,
		cell('Width', 2) + cell('Height', 1) + cell('PinX', 2) + cell('PinY', 2) + rectangle + extra,
	);
const remove = (shapeId: string) => ({ type: 'delete-shape' as const, pageId: '0', shapeId });
async function pageXml(bytes: Uint8Array): Promise<string> {
	const pkg = await VisioPackage.open(bytes);
	return new TextDecoder().decode(await pkg.readBytes('visio/pages/page1.xml'));
}
const open = (shapes: string) =>
	fixture({ pages: [{ id: '0', contents: `<Shapes>${shapes}</Shapes>` }] });

describe('deleting a shape that other cells read', () => {
	it('freezes the dependent cell to its last result and keeps the rest', async () => {
		const dependent = shape(
			'3',
			cell('Width', 2) +
				cell('Height', 1) +
				cell('PinX', 2) +
				cell('PinY', 2) +
				cell('LocPinX', 0.5, 'Sheet.1!Width*0.25') +
				rectangle,
		);
		const saved = await editVsdx(await open(leaf('1') + dependent), [remove('1')]);
		const xml = await pageXml(saved.bytes);
		expect(xml).not.toContain('Sheet.1!');
		expect(xml).toMatch(/<Cell N="LocPinX" V="0.5"\s*\/>/);
		expect(xml).not.toContain('ID="1"');
	});
	it('refuses a protected cell and a possibly dynamic reader', async () => {
		const guarded = shape(
			'3',
			cell('Width', 2) +
				cell('Height', 1) +
				cell('LocPinX', 1, 'GUARD(Sheet.1!Width*0.5)') +
				rectangle,
		);
		await expect(editVsdx(await open(leaf('1') + guarded), [remove('1')])).rejects.toThrow(
			/referenced/,
		);
		const dynamic = shape(
			'3',
			cell('Width', 2) +
				cell('Height', 1) +
				cell('LocPinX', 1, 'EVALCELL(&quot;Width&quot;)') +
				rectangle,
		);
		await expect(editVsdx(await open(leaf('1') + dynamic), [remove('1')])).rejects.toThrow(
			/dynamic/,
		);
	});
});
