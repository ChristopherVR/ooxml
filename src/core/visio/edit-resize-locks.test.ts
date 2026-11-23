import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, shape } from './test-fixtures';
import { visioResizeDrag } from './ui/shape-resize';
import { visioSnapResizeDrag } from './ui/resize-snap';
import { visioSizePositionCommand } from './ui/size-position';

const baseCells =
	cell('PinX', 3) +
	cell('PinY', 4) +
	cell('Width', 2) +
	cell('Height', 1) +
	cell('LocPinX', 1, 'Width*0.5') +
	cell('LocPinY', 0.5, 'Height*0.5');

const makeDoc = (extra = '', styles = '', attrs = '') =>
	fixture({
		document: styles,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', baseCells + extra + rectangle, attrs)}</Shapes>`,
			},
		],
	});

describe('Visio resize locks (LockWidth, LockHeight, LockAspect)', () => {
	it('resizes along height when LockWidth is active, including corner and Shift drags', async () => {
		const bytes = await makeDoc(cell('LockWidth', 1));
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(page.shapes[0]!.locks).toEqual({ width: true });

		// Corner drag moves both pointer axes, but width stays fixed and height resizes.
		const corner = visioResizeDrag(page, '1', 'ne', { x: 4, y: -4.5 }, { x: 5.5, y: -5.5 })!;
		expect(corner.command.width).toBeCloseTo(2, 10);
		expect(corner.command.height).toBeCloseTo(2, 10);
		const saved = await editVsdx(bytes, [corner.command]);
		const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(reopened.width).toBeCloseTo(2, 10);
		expect(reopened.height).toBeCloseTo(2, 10);

		// Shift+drag on north still keeps width fixed.
		const shiftN = visioResizeDrag(
			page,
			'1',
			'n',
			{ x: 3, y: -4.5 },
			{ x: 3, y: -5.5 },
			{
				proportional: true,
			},
		)!;
		expect(shiftN.command.width).toBeCloseTo(2, 10);
		expect(shiftN.command.height).toBeCloseTo(2, 10);

		// Dragging the locked east handle is a no-op rather than a refused edit.
		const east = visioResizeDrag(page, '1', 'e', { x: 4, y: -4 }, { x: 5, y: -4 })!;
		expect(east.command.width).toBeCloseTo(2, 10);
		expect(east.command.height).toBeCloseTo(1, 10);
		expect((await editVsdx(bytes, [east.command])).changedParts).toEqual([]);
	});

	it('resizes a rotated LockWidth shape along its local height axis', async () => {
		const bytes = await makeDoc(cell('LockWidth', 1) + cell('Angle', Math.PI / 6));
		const page = (await parseVsdx(bytes)).pages[0]!;
		const corner = visioResizeDrag(page, '1', 'ne', { x: 0, y: 0 }, { x: 0, y: -1 })!;
		expect(corner.command.width).toBeCloseTo(2, 10);
		expect(corner.command.height).toBeGreaterThan(1);
		const saved = await editVsdx(bytes, [corner.command]);
		const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(reopened.width).toBeCloseTo(2, 10);
		expect(reopened.height).toBeCloseTo(corner.command.height, 10);
	});

	it('resizes along width when LockHeight is active', async () => {
		const bytes = await makeDoc(cell('LockHeight', 1));
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(page.shapes[0]!.locks).toEqual({ height: true });

		const corner = visioResizeDrag(page, '1', 'se', { x: 4, y: -3.5 }, { x: 5, y: -2.5 })!;
		expect(corner.command.width).toBeCloseTo(3, 10);
		expect(corner.command.height).toBeCloseTo(1, 10);
		const saved = await editVsdx(bytes, [corner.command]);
		const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(reopened.width).toBeCloseTo(3, 10);
		expect(reopened.height).toBeCloseTo(1, 10);

		const north = visioResizeDrag(page, '1', 'n', { x: 3, y: -4.5 }, { x: 3, y: -6 })!;
		expect((await editVsdx(bytes, [north.command])).changedParts).toEqual([]);
	});

	it('scales proportionally on side handles and Size & Position when LockAspect is active', async () => {
		const bytes = await makeDoc(cell('LockAspect', 1));
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(page.shapes[0]!.locks).toEqual({ aspect: true });

		// East side handle without Shift keeps the 2:1 aspect ratio.
		const east = visioResizeDrag(page, '1', 'e', { x: 4, y: -4 }, { x: 6, y: -4 })!;
		expect(east.command.width).toBeCloseTo(4, 10);
		expect(east.command.height).toBeCloseTo(2, 10);
		const saved = await editVsdx(bytes, [east.command]);
		const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(reopened.width).toBeCloseTo(4, 10);
		expect(reopened.height).toBeCloseTo(2, 10);

		// Size & Position width edit also scales height proportionally.
		const sizeCmd = visioSizePositionCommand(page, '1', 'width', 3)!;
		expect(sizeCmd).toEqual([
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 3, height: 1.5 },
		]);
		const savedSize = await editVsdx(bytes, sizeCmd);
		const reopenedSize = (await parseVsdx(savedSize.bytes)).pages[0]!.shapes[0]!;
		expect(reopenedSize.width).toBeCloseTo(3, 10);
		expect(reopenedSize.height).toBeCloseTo(1.5, 10);
	});

	it.each([
		['LockWidth + LockHeight', cell('LockWidth', 1) + cell('LockHeight', 1)],
		['LockWidth + LockAspect', cell('LockWidth', 1) + cell('LockAspect', 1)],
		['LockHeight + LockAspect', cell('LockHeight', 1) + cell('LockAspect', 1)],
	])('prevents resizing on either axis when %s are both active', async (_label, locks) => {
		const bytes = await makeDoc(locks);
		const page = (await parseVsdx(bytes)).pages[0]!;
		const corner = visioResizeDrag(page, '1', 'ne', { x: 4, y: -4.5 }, { x: 6, y: -6.5 })!;
		expect(corner.command.width).toBeCloseTo(2, 10);
		expect(corner.command.height).toBeCloseTo(1, 10);
		expect((await editVsdx(bytes, [corner.command])).changedParts).toEqual([]);
	});

	it('honors style-inherited LockWidth and LockAspect per dimension and allows local overrides', async () => {
		const widthLockedStyle = `<StyleSheets><StyleSheet ID="1">${cell('LockWidth', 1)}</StyleSheet></StyleSheets>`;
		const bytes = await makeDoc('', widthLockedStyle, 'LineStyle="1" FillStyle="1" TextStyle="1"');
		const page = (await parseVsdx(bytes)).pages[0]!;
		expect(page.shapes[0]!.locks).toEqual({ width: true });

		// Dragging corner resizes height only and saves cleanly.
		const corner = visioResizeDrag(page, '1', 'ne', { x: 4, y: -4.5 }, { x: 6, y: -5.5 })!;
		const saved = await editVsdx(bytes, [corner.command]);
		const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(reopened.width).toBeCloseTo(2, 10);
		expect(reopened.height).toBeCloseTo(2, 10);

		// Directly requesting a width change is still refused.
		await expect(
			editVsdx(bytes, [{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 3, height: 1 }]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });

		// Local LockWidth="0" overrides inherited style LockWidth="1".
		const overridden = await makeDoc(
			cell('LockWidth', 0),
			widthLockedStyle,
			'LineStyle="1" FillStyle="1" TextStyle="1"',
		);
		const savedOverride = await editVsdx(overridden, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 3, height: 1 },
		]);
		expect((await parseVsdx(savedOverride.bytes)).pages[0]!.shapes[0]!.width).toBeCloseTo(3, 10);

		// Style-inherited LockAspect allows proportional resize and refuses non-proportional.
		const aspectStyle = `<StyleSheets><StyleSheet ID="1">${cell('LockAspect', 1)}</StyleSheet></StyleSheets>`;
		const aspectBytes = await makeDoc('', aspectStyle, 'LineStyle="1" FillStyle="1" TextStyle="1"');
		await expect(
			editVsdx(aspectBytes, [
				{
					type: 'resize-shape',
					pageId: '0',
					shapeId: '1',
					width: 4,
					height: 2,
					anchor: { x: 0, y: 0 },
				},
			]),
		).resolves.toHaveProperty('bytes');
		await expect(
			editVsdx(aspectBytes, [
				{
					type: 'resize-shape',
					pageId: '0',
					shapeId: '1',
					width: 4,
					height: 1.5,
					anchor: { x: 0, y: 0 },
				},
			]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});

	it('still refuses invalid or stale style-inherited resize lock cells', async () => {
		for (const bad of [
			'<Cell N="LockWidth" V="2"/>',
			'<Cell N="LockWidth" V="1" E="#VALUE!"/>',
			'<Cell N="LockWidth" V="1" F="0"/>',
		]) {
			const bytes = await makeDoc(
				'',
				`<StyleSheets><StyleSheet ID="1">${bad}</StyleSheet></StyleSheets>`,
				'LineStyle="1" FillStyle="1" TextStyle="1"',
			);
			await expect(
				editVsdx(bytes, [{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 2, height: 2 }]),
			).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		}
	});

	it('suppresses locked-axis snap guides during visioSnapResizeDrag', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents:
						`<Shapes>${shape('1', baseCells + cell('LockWidth', 1) + rectangle)}` +
						`${shape('2', cell('PinX', 5) + cell('PinY', 4) + cell('Width', 2) + cell('Height', 1) + rectangle)}</Shapes>`,
				},
			],
		});
		const page = (await parseVsdx(bytes)).pages[0]!;
		// East edge of shape 1 is at x=4, west edge of shape 2 is at x=4 (exact X alignment).
		const snapped = visioSnapResizeDrag(
			page,
			['1'],
			'ne',
			{ x: 4, y: -4.5 },
			{ x: 4.02, y: -5.5 },
			{
				snap: { shapes: true, guides: false, threshold: 0.1 },
			},
		)!;
		expect(snapped.lines.some((line) => line.axis === 'x')).toBe(false);
	});
});
