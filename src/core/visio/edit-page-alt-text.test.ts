import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture, cell, shape } from './test-fixtures';
import type { VisioPageAltTextEdit } from './edit-page-setup-commands';

const command: VisioPageAltTextEdit = {
	type: 'set-page-alt-text',
	pageId: '0',
	title: 'Process overview',
	description: 'Requests arrive\nthen proceed to approval.',
};

it('reads the native Visio page title and description fixture', async () => {
	const bytes = new Uint8Array(
		readFileSync(new URL('./__fixtures__/native-alt-text.vsdx', import.meta.url)),
	);
	expect((await parseVsdx(bytes)).pages[0]).toMatchObject({
		accessibilityTitle: 'Native page title',
		alternativeText: 'Native page description',
	});
});

it('sets, preserves and clears page alternatives independently of shapes and page identity', async () => {
	const input = await fixture();
	const result = await editVsdx(input, [command]);
	const page = (await parseVsdx(result.bytes)).pages[0]!;
	expect(page).toMatchObject({
		name: 'Page 1',
		accessibilityTitle: command.title,
		alternativeText: command.description,
	});
	expect(page.shapes[0]!.alternativeText).toBeUndefined();
	expect(result.changedParts).toEqual(['visio/pages/pages.xml']);
	const before = await VisioPackage.open(input),
		after = await VisioPackage.open(result.bytes);
	for (const path of before.paths().filter((path) => path !== 'visio/pages/pages.xml'))
		expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
	expect((await editVsdx(result.bytes, [command])).changedParts).toEqual([]);
	const cleared = await editVsdx(result.bytes, [{ ...command, title: '', description: '' }]);
	expect((await parseVsdx(cleared.bytes)).pages[0]).toMatchObject({
		accessibilityTitle: '',
		alternativeText: '',
	});
});

it('refuses formulas depending on page alternatives and validates text before mutation', async () => {
	const input = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', '2', 'LEN(ThePage!User.visAltText)'))}</Shapes>`,
			},
		],
	});
	await expect(editVsdx(input, [command])).rejects.toMatchObject({
		code: 'UNSUPPORTED_METADATA_EDIT',
	});
	await expect(
		editVsdx(await fixture(), [{ ...command, title: 'bad\nname' }]),
	).rejects.toMatchObject({ code: 'INVALID_EDIT_TEXT' });
});
