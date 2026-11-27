import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { assertVisioMasterIndependence } from './edit-master-scope';
import { VisioPackage } from './package';
import { cell, fixture, rectangle, shape } from './test-fixtures';

// The gallery's flowchart: 69 master instances (steps, decisions, stencil connectors glued with
// _XFTRIGGER, legend swatches) on a document with dozens of styles.
const sample = () =>
	new Uint8Array(
		readFileSync(
			new URL('../../../samples/visio/order-to-delivery-flowchart.vsdx', import.meta.url),
		),
	);
const move = (shapeId: string, x: number, y: number) =>
	({ type: 'move-shape', pageId: '0', shapeId, x, y }) as const;

describe('master proof on the order-to-delivery sample', () => {
	it('moves a local swimlane without exceeding the effective-cell construction budget', async () => {
		// Shape 6 is the Warehouse lane, 1 the title bar: neither is a master instance, but every
		// instance on the page is proven independent of them. Each master was rebuilt (styles,
		// markup, template cells) for every instance, about 190,000 units against a 100,000 budget.
		for (const command of [move('6', 7.45, 4.95), move('1', 7, 8.4)])
			await expect(editVsdx(sample(), [command])).resolves.toHaveProperty('bytes');
	});
});

describe('a master transform trigger on a page shape', () => {
	const dimensions = cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1);
	const local = (id: string) => shape(id, dimensions + rectangle);
	// A stencil instance whose formula reads another shape's transform event, as glue does.
	const glued = shape('2', cell('BegTrigger', 2, '_XFTRIGGER(Sheet.1!EventXFMod)'), 'Master="7"');
	const group = (contents: string) =>
		shape('4', dimensions + `<Shapes>${contents}</Shapes>`, 'Type="Group"');
	const drawing = (contents: string) =>
		fixture({
			masters: [{ id: '7', shapes: shape('10', dimensions + rectangle) }],
			pages: [{ id: '0', contents: `<Shapes>${contents}</Shapes>` }],
		});
	const prove = async (bytes: Uint8Array, shapeId: string) => {
		const pkg = await VisioPackage.open(bytes);
		const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
		return assertVisioMasterIndependence(
			pkg,
			new Map([['0', root]]),
			[move(shapeId, 5, 5)],
			() => {},
		);
	};

	it('is independent of a move of another shape', async () => {
		await expect(prove(await drawing(local('1') + glued + local('3')), '3')).resolves.toBeDefined();
	});
	it('depends on a move of the shape it reads, or of a group around it', async () => {
		await expect(prove(await drawing(local('1') + glued), '1')).rejects.toThrow(
			/edited page shape/,
		);
		await expect(prove(await drawing(group(local('1')) + glued), '4')).rejects.toThrow(
			/edited page shape/,
		);
	});
});
