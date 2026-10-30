import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { snapshotFormatting } from './edit-formatting-commands';
import { VISIO_BEVEL_PRESETS, visioRotation3dMatrix } from './bevel-3d';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { attribute, children } from './sheet';
import { cell, fixture, shape, rectangle } from './test-fixtures';
import { themeFixture } from './theme-fixtures';
import { themeArrow } from './theme-line';

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const plain = (extra = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 2) + cell('Height', 1) + rectangle + extra)}</Shapes>`,
			},
		],
	});
const read = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
async function cells(bytes: Uint8Array): Promise<Map<string, Element>> {
	const root = await (await VisioPackage.open(bytes)).readXml('visio/pages/page1.xml');
	const node = children(children(root, 'Shapes')[0], 'Shape')[0]!;
	return new Map(children(node, 'Cell').map((item) => [attribute(item, 'N')!, item]));
}

describe('Bevel and 3-D rotation', () => {
	it('numbers the twelve gallery bevels from 1', () => {
		expect(VISIO_BEVEL_PRESETS.map((preset) => preset.type)).toEqual([
			1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12,
		]);
		expect(VISIO_BEVEL_PRESETS[4]).toMatchObject({ id: 'angle', type: 5 });
	});
	it('writes the Bevel cells for a preset and reads them back', async () => {
		const saved = await editVsdx(await plain(), [{ ...target, bevel: 'soft-round' }]);
		const written = await cells(saved.bytes);
		expect(attribute(written.get('BevelTopType'), 'V')).toBe('6');
		expect(attribute(written.get('BevelTopWidth'), 'U')).toBe('PT');
		expect(Number(attribute(written.get('BevelTopWidth'), 'V'))).toBeCloseTo(6 / 72, 6);
		expect((await read(saved.bytes)).style.bevel).toMatchObject({ type: 6 });
		const cleared = await editVsdx(saved.bytes, [{ ...target, bevel: 'none' }]);
		expect((await read(cleared.bytes)).style.bevel).toBeUndefined();
	});
	it('writes rotation angles in radians and reads them back', async () => {
		const saved = await editVsdx(await plain(), [{ ...target, rotation3d: 'perspective-left' }]);
		const written = await cells(saved.bytes);
		expect(attribute(written.get('RotationXAngle'), 'U')).toBe('DEG');
		expect(Number(attribute(written.get('RotationXAngle'), 'V'))).toBeCloseTo(
			(-35 * Math.PI) / 180,
			6,
		);
		const rotation = (await read(saved.bytes)).style.rotation3d!;
		expect(rotation.x).toBeCloseTo((-35 * Math.PI) / 180, 6);
		expect(rotation.perspective).toBeCloseTo(Math.PI / 3, 6);
		const document = await parseVsdx(saved.bytes);
		expect(() => assertViewableDocument(document)).not.toThrow();
		expect(copySnapshotScene(document).pages[0]!.shapes[0]!.style.rotation3d).toMatchObject({
			y: 0,
		});
	});
	it('refuses an unknown preset', () => {
		expect(() => snapshotFormatting({ ...target, bevel: 'bogus' })).toThrow();
		expect(() => snapshotFormatting({ ...target, rotation3d: 'bogus' })).toThrow();
	});
	it('projects rotations as measured against Visio 16', () => {
		const [a, b, c, d] = visioRotation3dMatrix({ x: 0, y: 0, z: Math.PI / 6 });
		expect([a, b, c, d]).toEqual([
			Math.cos(Math.PI / 6),
			Math.sin(Math.PI / 6),
			-Math.sin(Math.PI / 6),
			Math.cos(Math.PI / 6),
		]);
		const turned = visioRotation3dMatrix({ x: (40 * Math.PI) / 180, y: 0, z: 0 });
		expect(turned[0]).toBeCloseTo(Math.cos((40 * Math.PI) / 180));
		expect(turned[3]).toBe(1);
		const both = visioRotation3dMatrix({ x: (40 * Math.PI) / 180, y: (40 * Math.PI) / 180, z: 0 });
		// The right side drops: x moves y downwards.
		expect(both[1]).toBeLessThan(0);
	});
	it('does not read a themed bevel from the theme', async () => {
		const themed = await plain(cell('BevelTopType', 'Themed', 'THEMEVAL()'));
		expect((await read(themed)).style.bevel).toBeUndefined();
	});
});

describe('Theme arrowheads', () => {
	it('keeps a literal arrowhead and ignores a non-themed cell', async () => {
		const bytes = await themeFixture({ contents: cell('EndArrow', 4) });
		expect((await read(bytes)).style.endArrow).toBe(4);
		expect(themeArrow(new Map([['EndArrow', { value: '4' }]]), {} as never, 'End')).toBeUndefined();
	});
});
