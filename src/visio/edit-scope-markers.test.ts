import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit.js';
import { parseVsdx } from './parser.js';
import { VisioPackageError } from './package-common.js';
import { cell, fixture, rectangle, shape } from './test-fixtures.js';

const geometry = (id = '1', attrs = '', markers = false) =>
	shape(
		id,
		cell('PinX', 2, markers ? 'No Formula' : '') +
			cell('PinY', 3) +
			cell('Width', 2, markers ? 'No Formula' : '') +
			cell('Height', 1) +
			rectangle,
		attrs,
	);
const move: VisioEdit = { type: 'move-shape', pageId: '0', shapeId: '1', x: 4, y: 5 };
const resize: VisioEdit = { type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 2 };
const remove: VisioEdit = { type: 'delete-shape', pageId: '0', shapeId: '1' };
const source = (document = '', contents = geometry()) =>
	fixture({ document, pages: [{ id: '0', contents: `<Shapes>${contents}</Shapes>` }] });
async function refusal(bytes: Uint8Array, commands: VisioEdit[], code: string) {
	const before = bytes.slice();
	let error: unknown;
	try {
		await editVsdx(bytes, commands);
	} catch (caught) {
		error = caught;
	}
	expect(error, 'An unsupported transaction must reject before returning bytes').toBeInstanceOf(
		VisioPackageError,
	);
	expect(error).toMatchObject({ code });
	expect(bytes).toEqual(before);
}
async function pageMetadata(pageId: string, formula: string) {
	const bytes = await fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${geometry()}</Shapes>` },
			{ id: '1', contents: `<Shapes>${geometry()}</Shapes>` },
		],
	});
	const zip = await JSZip.loadAsync(bytes);
	const original = await zip.file('visio/pages/pages.xml')!.async('string');
	zip.file(
		'visio/pages/pages.xml',
		original.replace(
			new RegExp(`(<Page ID="${pageId}"[^>]*><PageSheet>)`),
			`$1${cell('UserCache', 2, formula)}`,
		),
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('geometry scope and No Formula cache markers', () => {
	it('changes an imported direct numeric cache and preserves its No Formula marker', async () => {
		const bytes = await source('', geometry('1', '', true));
		const result = await editVsdx(bytes, [move, resize]);
		const zip = await JSZip.loadAsync(result.bytes),
			text = await zip.file('visio/pages/page1.xml')!.async('string');
		expect(text).toMatch(/N="PinX"[^>]*V="4"[^>]*F="No Formula"/);
		expect(text).toMatch(/N="Width"[^>]*V="4"[^>]*F="No Formula"/);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.width).toBe(4);
	});
	it('allows irrelevant document marker and foreign namespace formula metadata unchanged', async () => {
		const bytes = await source(
			'<DocumentSheet>' +
				cell('UserCache', 17, 'No Formula') +
				'<foreign:Cell xmlns:foreign="urn:opaque" N="Whatever" V="0" F="INDIRECT(&quot;Width&quot;)"/></DocumentSheet>',
		);
		const result = await editVsdx(bytes, [move]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(result.bytes);
		expect(await after.file('visio/document.xml')!.async('uint8array')).toEqual(
			await before.file('visio/document.xml')!.async('uint8array'),
		);
	});
	it('allows an irrelevant PageSheet No Formula marker', async () => {
		await expect(editVsdx(await pageMetadata('0', 'No Formula'), [move])).resolves.toHaveProperty(
			'bytes',
		);
	});
	it.each(['INDIRECT(&quot;Width&quot;)', 'UNKNOWN(Sheet.1!Width)'])(
		'rejects relevant document formula %s atomically',
		async (formula) => {
			await refusal(
				await source(`<DocumentSheet>${cell('UserCache', 2, formula)}</DocumentSheet>`),
				[move, resize],
				'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
			);
		},
	);
	it('refuses non-page transitive references to an affected page cell', async () => {
		const dependent = shape(
			'2',
			cell('PinX', 4, 'Sheet.1!Width*2') +
				cell('PinY', 3) +
				cell('Width', 1) +
				cell('Height', 1) +
				rectangle,
		);
		await refusal(
			await source(
				`<DocumentSheet>${cell('UserCache', 4, 'Sheet.2!PinX')}</DocumentSheet>`,
				geometry() + dependent,
			),
			[resize],
			'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
		);
	});
	it('treats master shape IDs and page shape IDs as separate scopes', async () => {
		const bytes = await fixture({
			masters: [
				{
					id: '7',
					shapes: shape(
						'1',
						cell('PinX', 4, 'Sheet.1!Width*2') + cell('Width', 2) + cell('Height', 1) + rectangle,
					),
				},
			],
			pages: [{ id: '0', contents: `<Shapes>${geometry()}</Shapes>` }],
		});
		const result = await editVsdx(bytes, [resize]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(result.bytes);
		expect(await after.file('visio/masters/master1.xml')!.async('uint8array')).toEqual(
			await before.file('visio/masters/master1.xml')!.async('uint8array'),
		);
	});
	it('does not mistake master Connect IDs for page references when deleting a local shape', async () => {
		const bytes = await fixture({
			masters: [{ id: '7', shapes: geometry() }],
			pages: [{ id: '0', contents: `<Shapes>${geometry()}</Shapes>` }],
		});
		const zip = await JSZip.loadAsync(bytes),
			xml = await zip.file('visio/masters/master1.xml')!.async('string');
		zip.file(
			'visio/masters/master1.xml',
			xml.replace(
				'</MasterContents>',
				'<Connects><Connect FromSheet="1" ToSheet="1" FromCell="BeginX" ToCell="PinX"/></Connects></MasterContents>',
			),
		);
		const result = await editVsdx(await zip.generateAsync({ type: 'uint8array' }), [remove]);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes).toHaveLength(0);
	});
	it('refuses relevant same-page PageSheet metadata caches', async () => {
		await refusal(
			await pageMetadata('0', 'Sheet.1!Width'),
			[resize],
			'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
		);
	});
	it('allows another page PageSheet static reference with the same numeric shape ID', async () => {
		await expect(
			editVsdx(await pageMetadata('1', 'Sheet.1!Width'), [resize]),
		).resolves.toHaveProperty('bytes');
	});
	it('refuses stale inherited style geometry formulas affected by target resize', async () => {
		const document = `<StyleSheets><StyleSheet ID="42">${cell('TxtWidth', 4, 'Width*2')}</StyleSheet></StyleSheets>`;
		await refusal(
			await source(document, geometry('1', 'TextStyle="42"')),
			[resize],
			'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
		);
	});
	it('keeps Inh, master targets and protected redirected formulas refused', async () => {
		await refusal(
			await source('', geometry('1', 'Master="7"')),
			[move],
			'UNSUPPORTED_GEOMETRY_EDIT',
		);
		const inherited = shape(
			'1',
			cell('PinX', 2, 'Inh') + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle,
		);
		await expect(editVsdx(await source('', inherited), [move])).rejects.toThrow();
		const guarded = shape(
			'1',
			cell('PinX', 2, 'SETATREF(User.X)') +
				cell('PinY', 3) +
				cell('Width', 2) +
				cell('Height', 1) +
				rectangle,
		);
		await refusal(await source('', guarded), [move], 'EDIT_PROTECTED_CELL');
	});
});
