import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture, shape } from './test-fixtures';

const PAGE = 'visio/pages/page1.xml';
const c = (name: string, value: string | number, formula = '', unit = '') =>
	`<Cell N="${name}" V="${value}"${unit ? ` U="${unit}"` : ''}${formula ? ` F="${formula}"` : ''}/>`;
const row = (index: number, type: string, cells: string) =>
	`<Row T="${type}" IX="${index}">${cells}</Row>`;
const action = (name: string, cells: string) => `<Row N="${name}">${cells}</Row>`;

/** A callout-like master: a box that hides and shows itself and a reset that needs Visio. */
const master = (actions = '') =>
	shape(
		'6',
		c('PinX', 2) +
			c('PinY', 2) +
			c('Width', 2) +
			c('Height', 1) +
			c('LocPinX', 1, 'Width*0.5') +
			c('LocPinY', 0.5, 'Height*0.5') +
			`<Section N="Actions">${
				actions ||
				action(
					'Row_1',
					c(
						'Menu',
						'Show Bo&amp;x',
						'IF(Geometry2.NoShow=1,&quot;Show Bo&amp;x&quot;,&quot;Hide Bo&amp;x&quot;)',
					) +
						c('Action', 0, 'SETF(GetRef(Geometry2.NoShow),NOT(Geometry2.NoShow))') +
						c('Checked', 0) +
						c('Disabled', 0) +
						c('BeginGroup', 1),
				) +
					action(
						'Row_2',
						c('Menu', 'Open Help') +
							c('Action', 0, 'OPENFILE(&quot;help.chm&quot;)') +
							c('Disabled', 0),
					) +
					action(
						'Row_3',
						c('Menu', 'Locked') +
							c('Action', 0, 'SETF(GetRef(Geometry2.NoShow),1)') +
							c('Disabled', 1),
					)
			}</Section>` +
			`<Section N="Geometry" IX="0">${c('NoFill', 0)}${row(1, 'MoveTo', c('X', 0) + c('Y', 0))}${row(2, 'LineTo', c('X', 2, 'Width*1') + c('Y', 0))}</Section>` +
			`<Section N="Geometry" IX="1">${c('NoFill', 1)}${c('NoShow', 1)}${row(1, 'MoveTo', c('X', 0) + c('Y', 0))}${row(2, 'LineTo', c('X', 1, 'Width*0.5') + c('Y', 1))}</Section>`,
		'Type="Shape"',
	);
const instance = `<Shape ID="1" Type="Shape" Master="2">${c('PinX', 2)}${c('PinY', 6)}</Shape>`;
const source = (definition = master()) =>
	fixture({
		masters: [{ id: '2', shapes: definition }],
		pages: [{ id: '0', contents: `<Shapes>${instance}</Shapes>` }],
	});
const run = (name: string): VisioEdit => ({
	type: 'run-action',
	pageId: '0',
	shapeId: '1',
	row: name,
});
const refusal = async (bytes: Uint8Array, edit: VisioEdit) =>
	(
		await editVsdx(bytes, [edit]).then(
			() => undefined,
			(caught: { code?: string }) => caught,
		)
	)?.code;

describe('actions in the model', () => {
	it('lists the menu rows with their state and whether they can run', async () => {
		const [box, help, locked] = (await parseVsdx(await source())).pages[0]!.shapes[0]!.actions!;
		expect(box).toMatchObject({
			row: 'Row_1',
			menu: 'Show Box',
			checked: false,
			disabled: false,
			beginGroup: true,
			run: { kind: 'setf', targets: ['Geometry2.NoShow'] },
		});
		expect(help!.run).toMatchObject({ kind: 'unsupported' });
		expect((help!.run as { reason: string }).reason).toMatch(/OPENFILE/);
		expect(locked!.disabled).toBe(true);
	});
});

describe('running an action', () => {
	it('sets the cell and refreshes the menu text, as Visio saves it', async () => {
		const saved = await editVsdx(await source(), [run('Row_1')]);
		const page = await (await JSZip.loadAsync(saved.bytes)).file(PAGE)!.async('string');
		expect(page).toContain('<Section N="Geometry" IX="1"><Cell N="NoShow" V="0"/></Section>');
		expect(page).toContain(
			'<Section N="Actions"><Row N="Row_1"><Cell N="Menu" V="Hide Bo&amp;x" F="Inh"/></Row></Section>',
		);
		const model = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(model.actions![0]!.menu).toBe('Hide Box');
		// Running it again puts both back.
		const again = await editVsdx(saved.bytes, [run('Row_1')]);
		expect((await parseVsdx(again.bytes)).pages[0]!.shapes[0]!.actions![0]!.menu).toBe('Show Box');
	});

	it('refuses actions that are disabled, outside the ShapeSheet, or missing', async () => {
		const bytes = await source();
		expect(await refusal(bytes, run('Row_3'))).toBe('EDIT_PROTECTED_CELL');
		expect(await refusal(bytes, run('Row_2'))).toBe('UNSUPPORTED_ACTION');
		expect(await refusal(bytes, run('Row_9'))).toBe('EDIT_TARGET_NOT_FOUND');
	});
});
