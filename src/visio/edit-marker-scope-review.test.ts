import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit.js';
import { cell, fixture, rectangle, shape } from './test-fixtures.js';

const dimensions = cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1);
const local = (extra = '', attrs = '') => shape('1', dimensions + rectangle + extra, attrs);
const move = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 4, y: 5 };
const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 4, height: 2 };
const source = (contents = local(), document = '') =>
	fixture({ document, pages: [{ id: '0', contents: `<Shapes>${contents}</Shapes>` }] });
const refused = (promise: Promise<unknown>) =>
	expect(promise.then(() => 'saved')).rejects.toThrow();

describe('independent marker and inherited scope review', () => {
	it('admits explicit inert OneD=0 cache marker for an existing local shape', async () => {
		await expect(
			editVsdx(await source(local(cell('OneD', 0, 'No Formula'))), [move]),
		).resolves.toHaveProperty('bytes');
	});
	it.each(['OPENTEXTWIN()', 'DEFAULTEVENT()'])(
		'admits the literal independent page event %s',
		async (formula) => {
			await expect(
				editVsdx(await source(local(cell('EventDblClick', 0, formula))), [move]),
			).resolves.toHaveProperty('bytes');
		},
	);
	it('refuses geometry-dependent event expressions instead of treating them as inert', async () => {
		await refused(
			editVsdx(await source(local(cell('EventDblClick', 0, 'OPENTEXTWIN(Width)'))), [resize]),
		);
	});
	it('refuses numeric marker cells carrying error metadata', async () => {
		const bytes = await source(
			local().replace('N="Width" V="2"', 'N="Width" V="2" F="No Formula" E="#VALUE!"'),
		);
		await refused(editVsdx(bytes, [resize]));
	});
	it('refuses dynamic dependencies in another page metadata sheet', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${local()}</Shapes>` },
				{ id: '1', contents: `<Shapes>${local()}</Shapes>` },
			],
		});
		const zip = await JSZip.loadAsync(bytes),
			path = 'visio/pages/pages.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace(
				/(<Page ID="1"[^>]*><PageSheet>)/,
				`$1${cell('UserCache', 2, 'EVALCELL(&quot;Pages[Page 1]!Sheet.1!Width&quot;)')}`,
			),
		);
		await refused(editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize]));
	});
	it('refuses inherited dynamic master formulas when the definition has a page instance', async () => {
		const bytes = await fixture({
			masters: [
				{
					id: '7',
					shapes: shape(
						'10',
						cell('Width', 2, 'EVALCELL(&quot;Pages[Page 1]!Sheet.1!Width&quot;)') +
							cell('Height', 1) +
							rectangle,
					),
				},
			],
			pages: [{ id: '0', contents: `<Shapes>${local()}${shape('2', '', 'Master="7"')}</Shapes>` }],
		});
		await refused(editVsdx(bytes, [resize]));
	});
	it('does not substitute a synthetic default for a constant inherited text dimension', async () => {
		const document = `<StyleSheets><StyleSheet ID="42">${cell('TxtWidth', 10, 'No Formula')}</StyleSheet></StyleSheets>`;
		const bytes = await source(
			local(cell('TxtPinX', 5, 'TxtWidth*0.5'), 'TextStyle="42"'),
			document,
		);
		let saved;
		try {
			saved = await editVsdx(bytes, [resize]);
		} catch {
			return;
		}
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(await zip.file('visio/pages/page1.xml')!.async('string')).toContain('N="TxtPinX" V="5"');
	});
	it('refuses geometry-dependent formulas in implicit default style zero', async () => {
		const document = `<StyleSheets><StyleSheet ID="0">${cell('TxtWidth', 4, 'Width*2')}</StyleSheet></StyleSheets>`;
		await refused(editVsdx(await source(local(), document), [resize]));
	});
	it('refuses an inherited custom User dependency chain', async () => {
		const document = `<StyleSheets><StyleSheet ID="42"><Section N="User"><Row N="first">${cell('Value', 2, 'Width')}</Row><Row N="second">${cell('Value', 4, 'User.first*2')}</Row></Section>${cell('TxtWidth', 4, 'User.second')}</StyleSheet></StyleSheets>`;
		await refused(
			editVsdx(
				await source(local(cell('TxtPinX', 2, 'TxtWidth*0.5'), 'TextStyle="42"'), document),
				[resize],
			),
		);
	});
	it('checks inherited default style formulas for shapes created by this transaction', async () => {
		const document = `<StyleSheets><StyleSheet ID="0">${cell('TxtWidth', 4, 'Width*2')}</StyleSheet></StyleSheets>`;
		await refused(
			editVsdx(await source(local(), document), [
				{ type: 'create-rectangle', pageId: '0', shapeId: '9', x: 4, y: 5, width: 4, height: 2 },
			]),
		);
	});
	it.each(['No Formula', '0', 'GUARD(0)'])(
		'admits proven inactive inherited scalar lock %s',
		async (formula) => {
			const document = `<StyleSheets><StyleSheet ID="42">${cell('LockMoveX', 0, formula)}</StyleSheet></StyleSheets>`;
			await expect(
				editVsdx(await source(local('', 'TextStyle="42"'), document), [move]),
			).resolves.toHaveProperty('bytes');
		},
	);
	it.each([
		'<Cell N="LockMoveX" V="0" F="0 rad"/>',
		'<Cell N="LockMoveX" V="0" F="No Formula" U="DL"/>',
		'<Cell N="LockMoveX" V="0" F="No Formula" E="#VALUE!"/>',
		'<Cell N="LockMoveX" V="0" F="Inh"/>',
		'<Cell N="LockMoveX" V="0" F="1"/>',
		'<Cell N="LockMoveX" V="0" F="Width*0"/>',
		'<Cell N="LockMoveX" V="0" F="SETATREF(User.target)"/>',
	])('refuses unknown, stale, redirected or dimensional inherited zero lock: %s', async (lock) => {
		const document = `<StyleSheets><StyleSheet ID="42">${lock}</StyleSheet></StyleSheets>`;
		await refused(editVsdx(await source(local('', 'TextStyle="42"'), document), [move]));
	});
});
