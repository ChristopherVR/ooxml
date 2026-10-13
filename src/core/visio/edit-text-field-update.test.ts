import { readFileSync } from 'node:fs';
import { parseXml, buildXml } from '../xml';
import { children, attribute } from './sheet';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { visioTextSource } from './ui/text-draft';
import type { VisioEdit } from './edit-commands';
import { fixture, shape } from './test-fixtures';

const native = readFileSync(new URL('./__fixtures__/field-edit-native.vsdx', import.meta.url));
async function command(
	bytes: Uint8Array,
	id = '1',
	formula = 'Width*2',
	format = '0.00',
): Promise<VisioEdit> {
	const page = (await parseVsdx(bytes)).pages[0]!,
		text = page.shapes.find((shape) => shape.id === id)!.text,
		field = text.fields![0]!;
	return {
		type: 'update-text-field',
		pageId: page.id,
		shapeId: id,
		fieldIndex: field.index!,
		formula,
		format,
		expectedText: visioTextSource(text).source,
		expectedCached: field.cached,
		expectedSourceKey: field.sourceKey!,
	};
}
const xml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
it.each([
	['1', 'Width*2', '0.00', 'Before 8.00 | 1 After'],
	['2', 'DOCCREATION()', '{{yyyy-MM-dd}}', 'date'],
	['3', 'TITLE()', '@', 'Before  | 1 After'],
])(
	'updates one existing native field without replacing neighbors: %s',
	async (id, formula, format, expected) => {
		const original = await xml(native),
			saved = await editVsdx(native, [await command(native, id, formula, format)]),
			reopened = (await parseVsdx(saved.bytes)).pages[0]!;
		if (expected === 'date')
			expect(reopened.shapes.find((shape) => shape.id === id)!.text.plainText).toMatch(
				/^Before \d{4}-\d{2}-\d{2} \| 1 After$/,
			);
		else expect(reopened.shapes.find((shape) => shape.id === id)!.text.plainText).toBe(expected);
		expect(reopened.shapes.find((shape) => shape.id === id)!.text.fields).toHaveLength(2);
		expect(reopened.shapes.find((shape) => shape.id === id)!.text.fields![0]!.index).toBe('0');
		const before = parseXml(original),
			after = parseXml(await xml(saved.bytes));
		const selected = (doc: Document) =>
			Array.from(doc.getElementsByTagName('Shape')).find(
				(shape) => shape.getAttribute('ID') === id,
			)!;
		const row = (shape: Element) =>
			Array.from(shape.getElementsByTagName('Section'))
				.find((section) => section.getAttribute('N') === 'Field')!
				.getElementsByTagName('Row')[1]!;
		expect(buildXml(row(selected(after)))).toBe(buildXml(row(selected(before))));
		for (const other of Array.from(before.getElementsByTagName('Shape')).filter(
			(shape) => attribute(shape, 'ID') !== id,
		))
			expect(
				buildXml(
					Array.from(after.getElementsByTagName('Shape')).find(
						(shape) => attribute(shape, 'ID') === attribute(other, 'ID'),
					)!,
				),
			).toBe(buildXml(other));
		const metadataCell = (doc: Document, name: string) =>
			children(
				children(
					children(selected(doc), 'Section').find((section) => attribute(section, 'N') === 'Field'),
					'Row',
				)[0],
				'Cell',
			).find((cell) => attribute(cell, 'N') === name)!;
		expect(attribute(metadataCell(after, 'Value'), 'U')).toBe(
			id === '1' ? 'IN' : id === '2' ? 'DATE' : 'STR',
		);
		expect(attribute(metadataCell(after, 'Format'), 'U')).toBe('STR');
		expect(attribute(metadataCell(after, 'Value'), 'F')).toBe(formula);
		expect(attribute(metadataCell(after, 'Format'), 'F')).toBe(`"${format.replaceAll('"', '""')}"`);
		for (const cellName of ['UICat', 'UICod', 'UIFmt', 'Calendar', 'ObjectKind'])
			expect(buildXml(metadataCell(after, cellName))).toBe(
				buildXml(metadataCell(before, cellName)),
			);
		expect(selected(after).getElementsByTagName('Text')[0]!.childNodes.length).toBe(
			selected(before).getElementsByTagName('Text')[0]!.childNodes.length,
		);
		for (const field of Array.from(selected(after).getElementsByTagName('fld')))
			expect(field.getAttribute('IX')).toBe(
				Array.from(selected(before).getElementsByTagName('fld'))[
					Array.from(selected(after).getElementsByTagName('fld')).indexOf(field)
				]!.getAttribute('IX'),
			);
		const originalZip = await JSZip.loadAsync(native),
			savedZip = await JSZip.loadAsync(saved.bytes);
		for (const path of Object.keys(originalZip.files).filter(
			(path) => !originalZip.files[path]!.dir && path !== 'visio/pages/page1.xml',
		))
			expect(await savedZip.file(path)!.async('uint8array')).toEqual(
				await originalZip.file(path)!.async('uint8array'),
			);
		const repeated = await editVsdx(saved.bytes, [await command(saved.bytes, id, formula, format)]);
		expect(repeated.bytes).toEqual(saved.bytes);
	},
);
it('refuses stale field identity, cache, source text and row keys atomically', async () => {
	const chosen = await command(native);
	for (const override of [
		{ fieldIndex: '9' },
		{ expectedCached: 'stale' },
		{ expectedText: 'stale' },
		{ expectedSourceKey: 'stale' },
	])
		await expect(editVsdx(native, [{ ...chosen, ...override } as VisioEdit])).rejects.toThrow();
	const edited = await editVsdx(native, [chosen]);
	await expect(editVsdx(edited.bytes, [chosen])).rejects.toThrow(/changed/);
});
it.each([
	'<Cell N="LockTextEdit" V="1"/>',
	'<Shapes><Shape ID="2" Type="Shape"/></Shapes>',
	'<ForeignData ForeignType="Bitmap"/>',
])('refuses protected, grouped and foreign field targets: %s', async (extra) => {
	const base = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes>' +
					shape(
						'1',
						extra +
							'<Section N="Field"><Row IX="0"><Cell N="Value" V="2" F="2"/></Row></Section><Text><fld IX="0">2</fld></Text>',
					) +
					'</Shapes>',
			},
		],
	});
	await expect(editVsdx(base, [await command(base)])).rejects.toThrow();
});

it.each([
	['<Cell N="Width" V="2" F="Field.0.Value"/>', ''],
	['<Cell N="Width" V="2" F="Fields.Value"/>', ''],
	['', '<Cell N="Value" V="2" F=" GUARD(2)"/>'],
	['', '<Cell N="Value" V="2" F="Inh"/>'],
	['', '<Cell N="Format" V="x" F="FIELDPICTURE(999)"/>'],
	['', '<Cell N="Calendar" V="1"/>'],
	['', '<Cell N="ObjectKind" V="1"/>'],
	['', '<Cell N="value" V="2"/>'],
])('refuses dependent/protected/unmeasured source field edits', async (extra, cells) => {
	const value = cells.includes('N="Value"') ? '' : '<Cell N="Value" V="2" F="2"/>';
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes>' +
					shape(
						'1',
						extra +
							'<Section N="Field"><Row IX="0">' +
							value +
							cells +
							'</Row></Section><Text><fld IX="0">2</fld> after</Text>',
					) +
					'</Shapes>',
			},
		],
	});
	await expect(editVsdx(bytes, [await command(bytes)])).rejects.toThrow();
});
it('refuses duplicate field markers and grouped local child targets', async () => {
	const contents =
		'<Section N="Field"><Row IX="0"><Cell N="Value" V="2" F="2"/></Row></Section><Text><fld IX="0">2</fld><fld IX="0">2</fld></Text>';
	const duplicate = await fixture({
		pages: [{ id: '0', contents: '<Shapes>' + shape('1', contents) + '</Shapes>' }],
	});
	await expect(editVsdx(duplicate, [await command(duplicate)])).rejects.toThrow(/unique/);
	const nested = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes><Shape ID="2" Type="Group"><Shapes>' +
					shape(
						'1',
						contents.replace('<fld IX="0">2</fld><fld IX="0">2</fld>', '<fld IX="0">2</fld>'),
					) +
					'</Shapes></Shape></Shapes>',
			},
		],
	});
	const page = (await parseVsdx(nested)).pages[0]!,
		text = page.shapes[0]!.children[0]!.text,
		field = text.fields![0]!;
	await expect(
		editVsdx(nested, [
			{
				type: 'update-text-field',
				pageId: '0',
				shapeId: '1',
				fieldIndex: '0',
				formula: '3',
				format: '0',
				expectedText: visioTextSource(text).source,
				expectedCached: field.cached,
				expectedSourceKey: field.sourceKey!,
			},
		]),
	).rejects.toThrow(/top-level/);
});

it('retains source formatting markers, unknown row attributes/cells/XML and literal neighbors', async () => {
	const zip = await JSZip.loadAsync(native),
		source = await zip.file('visio/pages/page1.xml')!.async('string'),
		doc = parseXml(source),
		target = Array.from(doc.getElementsByTagName('Shape')).find(
			(node) => attribute(node, 'ID') === '1',
		)!,
		row = children(
			children(target, 'Section').find((node) => attribute(node, 'N') === 'Field'),
			'Row',
		)[0]!;
	row.setAttribute('UnknownAttribute', 'keep');
	const unknown = doc.createElementNS('urn:field-test', 'x:Extension');
	unknown.setAttribute('keep', 'yes');
	row.appendChild(unknown);
	const mystery = doc.createElementNS(row.namespaceURI, 'Cell');
	mystery.setAttribute('N', 'Mystery');
	mystery.setAttribute('V', 'unchanged');
	row.appendChild(mystery);
	zip.file('visio/pages/page1.xml', buildXml(doc));
	const bytes = await zip.generateAsync({ type: 'uint8array' }),
		saved = await editVsdx(bytes, [await command(bytes)]),
		after = parseXml(await xml(saved.bytes));
	const updated = Array.from(after.getElementsByTagName('Shape')).find(
			(node) => attribute(node, 'ID') === '1',
		)!,
		updatedRow = children(
			children(updated, 'Section').find((node) => attribute(node, 'N') === 'Field'),
			'Row',
		)[0]!;
	expect(attribute(updatedRow, 'UnknownAttribute')).toBe('keep');
	expect(buildXml(updatedRow.getElementsByTagNameNS('urn:field-test', 'Extension')[0]!)).toBe(
		buildXml(unknown),
	);
	expect(
		buildXml(children(updatedRow, 'Cell').find((cell) => attribute(cell, 'N') === 'Mystery')!),
	).toBe(buildXml(mystery));
	const beforeNodes = Array.from(children(target, 'Text')[0]!.childNodes),
		afterNodes = Array.from(children(updated, 'Text')[0]!.childNodes);
	expect(afterNodes).toHaveLength(beforeNodes.length);
	for (const [index, node] of beforeNodes.entries())
		if (
			!(
				node.nodeType === 1 &&
				(node as Element).localName === 'fld' &&
				attribute(node as Element, 'IX') === '0'
			)
		)
			expect(buildXml(afterNodes[index]!)).toBe(buildXml(node));
});
it('refuses new field-row self references and dependents on newly added field cells', async () => {
	await expect(
		editVsdx(native, [await command(native, '1', 'Field.0.Value+1', '0')]),
	).rejects.toThrow(/cannot reference/);
	for (const dependency of ['Field.0.Format', 'Field.0.Type']) {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents:
						'<Shapes>' +
						shape(
							'1',
							'<Cell N="Width" V="2" F="' +
								dependency +
								'"/><Section N="Field"><Row IX="0"><Cell N="Value" V="2" F="2"/></Row></Section><Text><fld IX="0">2</fld></Text>',
						) +
						'</Shapes>',
				},
			],
		});
		await expect(editVsdx(bytes, [await command(bytes)])).rejects.toThrow();
	}
});

it('rejects invalid XML formula strings before authoring a field cache', async () => {
	for (const formula of ['"bad\u0001"', '"bad\ud800"'])
		await expect(
			editVsdx(native, [await command(native, '1', formula, '@')]),
		).rejects.toMatchObject({ code: 'INVALID_EDIT_TEXT' });
});

it('refuses a unique lowercase alias instead of creating a duplicate canonical cell', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes>' +
					shape(
						'1',
						'<Section N="Field"><Row IX="0"><Cell N="value" V="2" F="2"/></Row></Section><Text><fld IX="0">2</fld></Text>',
					) +
					'</Shapes>',
			},
		],
	});
	await expect(editVsdx(bytes, [await command(bytes)])).rejects.toThrow(/Noncanonical/);
});
