import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { checkVisioAlternativeText } from './accessibility';
import { cell, fixture, rectangle, shape } from './test-fixtures';
import type { VisioShapeAltTextEdit } from './edit-metadata-commands';

const command: VisioShapeAltTextEdit = {
	type: 'set-shape-alt-text',
	pageId: '0',
	shapeId: '1',
	title: 'Service status',
	description: 'Green rectangle:\nall services operational.',
};
const drawing = (extra = '') =>
	fixture({ pages: [{ id: '0', contents: `<Shapes>${shape('1', rectangle + extra)}</Shapes>` }] });
const row = (name: string, value: string, formula?: string) =>
	`<Row N="${name}">${cell('Value', value, formula)}</Row>`;
const first = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;

it('reads titles and descriptions from a real native Visio drawing', async () => {
	const bytes = new Uint8Array(
		readFileSync(new URL('./__fixtures__/native-alt-text.vsdx', import.meta.url)),
	);
	expect(await first(bytes)).toMatchObject({
		accessibilityTitle: 'Native title',
		alternativeText: 'Native description',
	});
});

it('round trips native title and multiline description independently of ScreenTips and other User rows', async () => {
	const input = await drawing(
		`<Section N="User">${row('Custom', 'keep')}</Section>${cell('Comment', 'Hover help')}`,
	);
	const result = await editVsdx(input, [command]);
	expect(await first(result.bytes)).toMatchObject({
		accessibilityTitle: command.title,
		alternativeText: command.description,
		screenTip: 'Hover help',
	});
	const pkg = await VisioPackage.open(result.bytes);
	const xml = new TextDecoder().decode(await pkg.readBytes('visio/pages/page1.xml'));
	expect(xml).toContain('N="Custom"');
	expect(xml).toContain('U="STR"');
	expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
	expect((await editVsdx(result.bytes, [command])).changedParts).toEqual([]);
	const cleared = await editVsdx(result.bytes, [{ ...command, title: '', description: '' }]);
	expect(await first(cleared.bytes)).toMatchObject({
		accessibilityTitle: '',
		alternativeText: '',
		screenTip: 'Hover help',
	});
});

it('reads inherited native alternative text and respects local overrides and deleted rows', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', `<Section N="User">${row('visAltText', 'Local title')}<Row N="visAltDescription" Del="1"/></Section>`, 'Master="2"')}</Shapes>`,
			},
		],
		masters: [
			{
				id: '2',
				shapes: shape(
					'1',
					rectangle +
						`<Section N="User">${row('visAltText', 'Master title')}${row('visAltDescription', 'Master description')}</Section>`,
				),
			},
		],
	});
	expect(await first(bytes)).toMatchObject({ accessibilityTitle: 'Local title' });
	expect((await first(bytes)).alternativeText).toBeUndefined();
});

it('overrides and clears inherited stencil alternatives without changing its master or other rows', async () => {
	const bytes = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', '', 'Master="2"')}</Shapes>` }],
		masters: [
			{
				id: '2',
				shapes: shape(
					'1',
					rectangle +
						`<Section N="User">${row('visAltText', 'Master title')}${row('visAltDescription', 'Master description')}${row('Custom', 'Keep')}</Section>`,
				),
			},
		],
	});
	const changed = await editVsdx(bytes, [command]);
	expect(await first(changed.bytes)).toMatchObject({
		accessibilityTitle: command.title,
		alternativeText: command.description,
	});
	const before = await VisioPackage.open(bytes),
		after = await VisioPackage.open(changed.bytes);
	for (const path of before.paths().filter((path) => path.includes('masters/')))
		expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
	const cleared = await editVsdx(changed.bytes, [{ ...command, title: '', description: '' }]);
	expect(await first(cleared.bytes)).toMatchObject({ accessibilityTitle: '', alternativeText: '' });
});

it('refuses inherited formula-driven alternatives and references in a master', async () => {
	for (const contents of [
		`<Section N="User">${row('visAltText', 'cached', 'Prop.Title')}</Section>`,
		cell('Width', '2', 'LEN(User.visAltDescription)'),
	]) {
		const bytes = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${shape('1', '', 'Master="2"')}</Shapes>` }],
			masters: [{ id: '2', shapes: shape('1', rectangle + contents) }],
		});
		await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
			code: 'UNSUPPORTED_METADATA_EDIT',
		});
	}
});

it('refuses formula dependencies and ambiguous or computed metadata without mutating the source', async () => {
	for (const extra of [
		`<Section N="User">${row('visAltText', 'cached', 'Prop.Title')}</Section>`,
		`<Section N="User">${row('visAltText', 'one')}${row('visAltText', 'two')}</Section>`,
		cell('Width', '2', 'LEN(User.visAltDescription)'),
	]) {
		const bytes = await drawing(extra);
		const original = new Uint8Array(bytes);
		await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
			code: 'UNSUPPORTED_METADATA_EDIT',
		});
		expect(bytes).toEqual(original);
	}
	await expect(
		editVsdx(await drawing(), [{ ...command, description: '\u0001' }]),
	).rejects.toMatchObject({ code: 'INVALID_EDIT_TEXT' });
});

it('checks unlabeled graphics, accepts visible text and native alternatives, and skips hidden graphics', async () => {
	const model = await parseVsdx(await drawing());
	expect(checkVisioAlternativeText(model)).toEqual([
		{
			pageId: '0',
			shapeId: '1',
			shapeName: model.pages[0]!.shapes[0]!.name,
			code: 'missing-alternative-text',
		},
	]);
	const described = await parseVsdx((await editVsdx(await drawing(), [command])).bytes);
	expect(checkVisioAlternativeText(described)).toEqual([]);
	model.pages[0]!.shapes[0]!.hidden = true;
	expect(checkVisioAlternativeText(model)).toEqual([]);
});
