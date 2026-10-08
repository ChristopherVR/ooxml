import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { fixture, cell, shape } from './test-fixtures';

const command = { type: 'set-page-size' as const, pageId: '0', width: 6.25, height: 4.75 };
const pageCells =
	cell('DrawingSizeType', 3) +
	cell('DrawingResizeType', 0) +
	cell('PageScale', 1) +
	cell('DrawingScale', 1);
it.each(['Page.1', '[x]', 'A+B', 'Foo(1)'])(
	'keeps literal other-page names independent: %s',
	async (name) => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cell('PinX', 1, `Pages[${name}]!ThePage!PageWidth`))}</Shapes>`,
					pageCells,
				},
				{ id: '1', contents: '<Shapes/>', attributes: `NameU="${name}"`, pageCells },
			],
		});
		expect((await editVsdx(bytes, [command])).changedParts).toEqual(['visio/pages/pages.xml']);
	},
);
it.each(['Page.1', '[x]', 'A+B', 'Foo(1)'])(
	'does not mask literal target-page names: %s',
	async (name) => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: '<Shapes/>', attributes: `NameU="${name}"`, pageCells },
				{
					id: '1',
					contents: `<Shapes>${shape('1', cell('Width', 1, `Pages[${name}]!ThePage!PageWidth`))}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
			code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY',
		});
	},
);
it('does not mask case-folded aliases shared by target and independent page names', async () => {
	const bytes = await fixture({
		pages: [
			{ id: '0', contents: '<Shapes/>', attributes: 'NameU="SHARED"', pageCells },
			{
				id: '1',
				contents: `<Shapes>${shape('1', cell('Width', 1, 'Pages[shared]!ThePage!PageWidth'))}</Shapes>`,
				attributes: 'NameU="shared"',
			},
		],
	});
	await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
		code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY',
	});
});
it.each(['_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)', '_XFTRIGGER(Sheet.1!EventXFMod)'])(
	'refuses native glue formulas outside connected direct shape cells: %s',
	async (formula) => {
		for (const contents of [
			`<Cell N="BeginX" V="0" F="${formula}"/>`,
			`<Shapes>${shape('1', cell('BeginX', 0, formula))}</Shapes>`,
			`<Shapes>${shape('1', `<Section N="User"><Row>${cell('BegTrigger', 0, formula)}</Row></Section>`)}</Shapes><Connects><Connect FromSheet="1" ToSheet="1"/></Connects>`,
		]) {
			const bytes = await fixture({ pages: [{ id: '0', contents, pageCells }] });
			await expect(editVsdx(bytes, [command])).rejects.toMatchObject({
				code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY',
			});
		}
	},
);
it('leaves literal formula strings independent of page references and unknown function names', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', '<Section N="User"><Row N="Label"><Cell N="Value" V="literal" U="STR" F="&quot;ThePage!PageWidth UNKNOWNFUNCTION()&quot;"/></Row></Section>')}</Shapes>`,
				pageCells,
			},
		],
	});
	expect((await editVsdx(bytes, [command])).changedParts).toEqual(['visio/pages/pages.xml']);
});
