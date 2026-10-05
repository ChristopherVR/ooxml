import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit.js';
import { cell, fixture, rectangle, shape } from './test-fixtures.js';

const locks = ['LockMoveX', 'LockMoveY', 'LockWidth', 'LockHeight', 'LockAspect', 'LockDelete'];
const zero = locks.map((name) => cell(name, 0, 'No Formula')).join('');
const inh = locks.map((name) => cell(name, 0, 'Inh')).join('');
const parents = 'LineStyle="0" FillStyle="0" TextStyle="0"';
const target = (extra = '') =>
	shape(
		'1',
		cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + extra + rectangle,
		'LineStyle="7" FillStyle="7" TextStyle="7"',
	);
const move = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 4, y: 5 };
const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 4, height: 2 };
const document = (parent = zero, attrs = parents) =>
	`<StyleSheets><StyleSheet ID="0">${parent}</StyleSheet><StyleSheet ID="7" ${attrs}>${inh}</StyleSheet></StyleSheets>`;
const source = (styles = document(), extra = '') =>
	fixture({
		document: styles,
		pages: [{ id: '0', contents: `<Shapes>${target(extra)}</Shapes>` }],
	});
const refused = (promise: Promise<unknown>) =>
	expect(promise.then(() => 'saved')).rejects.toThrow();

describe('independent inherited protection ancestry review', () => {
	it('admits move and resize through explicit inactive style-parent ancestry', async () => {
		await expect(editVsdx(await source(), [move, resize])).resolves.toHaveProperty('bytes');
	});
	it.each([
		['active', cell('LockMoveX', 1)],
		['error', '<Cell N="LockMoveX" V="0" F="No Formula" E="#VALUE!"/>'],
		['unknown', cell('LockMoveX', 0, 'UNKNOWN(0)')],
		['dynamic', cell('LockMoveX', 0, 'SETATREF(User.value)')],
		['dimensional', cell('LockMoveX', 0, '0 rad')],
		['stale', cell('LockMoveX', 0, '1')],
	])('refuses a %s inherited parent lock', async (_name, lock) => {
		await refused(editVsdx(await source(document(lock)), [move]));
	});
	it('refuses an unresolved explicit parent', async () => {
		await refused(
			editVsdx(
				await source(
					`<StyleSheets><StyleSheet ID="7" ${parents}>${inh}</StyleSheet></StyleSheets>`,
				),
				[move],
			),
		);
	});
	it('refuses a parent ancestry cycle', async () => {
		const styles = `<StyleSheets><StyleSheet ID="7" LineStyle="8" FillStyle="8" TextStyle="8">${inh}</StyleSheet><StyleSheet ID="8" LineStyle="7" FillStyle="7" TextStyle="7">${inh}</StyleSheet></StyleSheets>`;
		await refused(editVsdx(await source(styles), [move]));
	});
	it('does not assume a parent when Inh has no explicit ancestry', async () => {
		await refused(editVsdx(await source(document(zero, '')), [move]));
	});
	it('keeps local inherited lock cells conservative', async () => {
		await refused(editVsdx(await source(document(), cell('LockMoveX', 0, 'Inh')), [move]));
	});
	it('refuses inconsistent nonzero child Inh caches', async () => {
		await refused(
			editVsdx(
				await source(
					document().replace(
						'<Cell N="LockMoveX" V="0" F="Inh"/>',
						'<Cell N="LockMoveX" V="1" F="Inh"/>',
					),
				),
				[move],
			),
		);
	});
	it('deletes an unreferenced shape whose own inherited style formula disappears with the instance', async () => {
		const styles = document().replace(
			`${inh}</StyleSheet>`,
			`${inh}${cell('LineWeight', 0.04, 'Width*0.02')}</StyleSheet>`,
		);
		await expect(
			editVsdx(await source(styles), [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]),
		).resolves.toHaveProperty('bytes');
	});
	it('proves a guarded inactive scalar parent without overwriting protection cells', async () => {
		const bytes = await source(document(zero.replace('F="No Formula"', 'F="GUARD(0)"')));
		const saved = await editVsdx(bytes, [move]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		expect(await after.file('visio/document.xml')!.async('uint8array')).toEqual(
			await before.file('visio/document.xml')!.async('uint8array'),
		);
	});
	it('refuses deletion referenced by another surviving instance inherited style', async () => {
		const styles = `<StyleSheets><StyleSheet ID="42">${cell('LineWeight', 0.04, 'Sheet.1!Width*0.02')}</StyleSheet></StyleSheets>`;
		const bytes = await fixture({
			document: styles,
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle)}${shape('2', cell('PinX', 5) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle, 'LineStyle="42"')}</Shapes>`,
				},
			],
		});
		await refused(editVsdx(bytes, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]));
	});
});
