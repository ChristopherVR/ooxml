import { describe, expect, it, vi } from 'vitest';
import { parseXml } from '../xml/index.js';
import { parseVsdx } from './index.js';
import { createMetadataBudget, shapeMetadata } from './shape-metadata.js';
import { mergeSheets, readSheet } from './sheet.js';
import { cell, fixture, section, shape, xml } from './test-fixtures.js';

const sheet = (data: string) => readSheet(parseXml(xml('Shape', data)).documentElement);
const named = (name: string, data: string, attributes = '') =>
	`<Row N="${name}" ${attributes}>${data}</Row>`;
const prop = (data: string, name = 'Property') => section(name, data);
const link = (data: string) => section('Hyperlink', named('Link', data));
const read = (data: string) => {
	const reports: string[] = [];
	return { ...shapeMetadata(sheet(data), (code) => reports.push(code)), reports };
};

describe('cached Visio shape data', () => {
	it.each([
		[0, 'hello &amp; goodbye', 'string', 'hello & goodbye'],
		[1, 'Option A', 'fixed-list', 'Option A'],
		[2, '-12.25e2', 'number', -1225],
		[3, 'TRUE', 'boolean', true],
		[3, '-1', 'boolean', true],
		[3, '0', 'boolean', false],
		[4, 'Custom option', 'variable-list', 'Custom option'],
		[5, '45291.5', 'date', 45291.5],
		[6, '0.25', 'duration', 0.25],
		[7, '23.45', 'currency', 23.45],
	] as const)(
		'preserves Type %s as %s without display formatting',
		(type, rawValue, valueKind, value) => {
			const result = read(prop(named('Asset', cell('Type', type) + cell('Value', rawValue))));
			expect(result.shapeData[0]).toMatchObject({
				id: 'Asset',
				name: 'Asset',
				type,
				valueKind,
				value,
			});
			expect(result.reports).toEqual([]);
		},
	);
	it('preserves labels, prompt, raw values, units, list format and visibility as inert text', () => {
		const result = read(
			prop(
				named(
					'Asset',
					cell('Type', 1) +
						'<Cell N="Value" V="&lt;script&gt;" U="STR"/>' +
						cell('Label', 'Asset label') +
						cell('Prompt', 'Read &amp; review') +
						cell('Format', 'A;B;&lt;script&gt;') +
						cell('Invisible', 1) +
						cell('SortKey', '01'),
				),
			),
		);
		expect(result.shapeData).toEqual([
			{
				id: 'Asset',
				name: 'Asset',
				label: 'Asset label',
				type: 1,
				rawType: '1',
				valueKind: 'fixed-list',
				rawValue: '<script>',
				value: '<script>',
				unit: 'STR',
				prompt: 'Read & review',
				format: 'A;B;<script>',
				invisible: true,
				sortKey: '01',
			},
		]);
	});
	it('uses default string type and retains an explicitly empty cached value', () => {
		expect(read(prop(named('Empty', cell('Value', '')))).shapeData[0]).toMatchObject({
			type: 0,
			valueKind: 'string',
			rawValue: '',
			value: '',
		});
	});
	it.each([
		[2, 'NaN'],
		[2, '1e999'],
		[2, '1.2.3'],
		[2, ''],
		[2, '0x10'],
		[3, 'yes'],
		[5, '31/12/2024'],
		[6, '4 hours'],
		[7, '$5'],
		[99, 'opaque'],
		[-2, 'opaque'],
		[8, 'opaque'],
		['bad', 'opaque'],
		[1.5, 'opaque'],
	])('retains unsupported type/value %s %s without inventing values', (type, value) => {
		const result = read(prop(named('Unsupported', cell('Type', type) + cell('Value', value))));
		expect(result.shapeData[0]).toMatchObject({
			rawType: String(type),
			rawValue: value,
			valueKind: 'unparsed',
		});
		expect(result.shapeData[0]).not.toHaveProperty('value');
		expect(result.reports).toContain('unsupported-shape-data-value');
	});
	it('retains stale cached text and errors without treating it as current typed data', () => {
		const result = read(
			prop(named('Error', cell('Type', 2) + '<Cell N="Value" V="12" E="#VALUE!" F="1/0"/>')),
		);
		expect(result.shapeData[0]).toMatchObject({
			rawValue: '12',
			error: '#VALUE!',
			valueKind: 'unparsed',
		});
		expect(result.shapeData[0]).not.toHaveProperty('value');
		expect(result.reports).toContain('metadata-cell-error');
	});
	it('does not execute formula-only values, infer type from a formula, or fabricate booleans', () => {
		const result = read(
			prop(
				named(
					'Formula',
					'<Cell N="Value" F="RUNADDON(&quot;no&quot;)"/><Cell N="Type" F="2"/>' +
						cell('Invisible', 'yes'),
				),
			),
		);
		expect(result.shapeData[0]).toMatchObject({ type: -1, valueKind: 'unparsed' });
		expect(result.shapeData[0]).not.toHaveProperty('value');
		expect(result.shapeData[0]).not.toHaveProperty('invisible');
		expect(result.reports).toContain('missing-metadata-cache');
		expect(result.reports).toContain('invalid-metadata-boolean');
	});
	it('honors merged master cells, deleted rows/sections and explicit empty local overrides', () => {
		const base = sheet(
			prop(
				named('Keep', cell('Label', 'From master') + cell('Type', 2) + cell('Value', 4)) +
					named('Remove', cell('Value', 'gone')),
			),
		);
		const local = sheet(
			prop(
				named('Keep', '<Cell N="Type" F="Inh"/>' + cell('Value', 9) + cell('Prompt', '')) +
					named('Remove', '', 'Del="1"'),
			),
		);
		const result = shapeMetadata(mergeSheets(base, local), () => {});
		expect(result.shapeData).toHaveLength(1);
		expect(result.shapeData[0]).toMatchObject({
			name: 'Keep',
			label: 'From master',
			type: 2,
			value: 9,
			prompt: '',
		});
		expect(
			shapeMetadata(mergeSheets(base, sheet('<Section N="Property" Del="1"/>')), () => {})
				.shapeData,
		).toEqual([]);
	});
	it('preserves inherited units on cached Inh overrides and keeps local errors', () => {
		const base = sheet(prop(named('Size', cell('Type', 2) + '<Cell N="Value" V="1" U="IN"/>')));
		const local = sheet(prop(named('Size', '<Cell N="Value" V="3" F="Inh"/>')));
		expect(shapeMetadata(mergeSheets(base, local), () => {}).shapeData[0]).toMatchObject({
			value: 3,
			unit: 'IN',
		});
		const error = sheet(prop(named('Size', '<Cell N="Value" F="Inh" E="#VALUE!"/>')));
		const result = shapeMetadata(mergeSheets(base, error), () => {}).shapeData[0];
		expect(result).toMatchObject({ unit: 'IN', error: '#VALUE!', valueKind: 'unparsed' });
		expect(result).not.toHaveProperty('value');
	});
	it('accepts the Prop shorthand without treating other sections as shape data', () => {
		expect(
			read(
				prop(named('Alias', cell('Value', 'yes')), 'Prop') +
					section('User', named('No', cell('Value', 'no'))),
			).shapeData,
		).toHaveLength(1);
	});
});

describe('inert Visio hyperlinks', () => {
	it.each([
		'https://example.com/path?q=1#part',
		'HTTP://example.com',
		'https://例子.测试/路径',
		'mailto:person@example.com?subject=Hello%20world',
	])('separates a validated address from raw cached text: %s', (address) => {
		const result = read(
			link(
				cell('Address', address) +
					cell('Description', '&lt;Open&gt;') +
					cell('Default', 1) +
					cell('NewWindow', 0) +
					cell('Invisible', 0),
			),
		);
		expect(result.hyperlinks[0]).toMatchObject({
			address,
			description: '<Open>',
			default: true,
			newWindow: false,
			invisible: false,
			target: { kind: 'external', href: new URL(address).href },
		});
		expect(result.reports).toEqual([]);
	});
	it.each([
		'javascript:alert(1)',
		'JaVaScRiPt:alert(1)',
		'data:text/html,hello',
		'file:///etc/passwd',
		'ftp://example.com',
		'tel:+10000000000',
		'../Drawing.vsdx',
		'//example.com',
		'\\\\server\\share',
		'https:example.com',
		'https:/example.com',
		'https:///example.com',
		'https://@example.com',
		'https://',
		'https://user:password@example.com',
		'https://user@example.com',
		'https://%75ser@example.com',
		'https://example.com/%',
		'https://example.com/%GG',
		'https://example.com/%C0%AF',
		'https://example.com/%E0%A4%A',
		'https://example.com/%0a',
		'https://example.com/%7f',
		'https://example.com/%C2%85',
		'https://example.com/%5Cpath',
		' https://example.com',
		'https://example.com/path with space',
		'https://example.com\\@evil.test',
		'mailto:',
		'mailto:not-an-address',
		'mailto://person@example.com',
		'mailto:person@example.com?subject=%0D%0ABcc:other@example.com',
	])('retains rejected address as inert data: %s', (address) => {
		const result = read(link(cell('Address', address)));
		expect(result.hyperlinks[0]).toMatchObject({
			address,
			target: { kind: 'unresolved', reason: 'unsafe-or-unsupported-address' },
		});
		expect(result.hyperlinks[0]!.target).not.toHaveProperty('href');
		expect(result.reports).toContain('unresolved-hyperlink');
	});
	it.each([
		'https://example.com/\npath',
		'https://example.com/\u0000path',
		'https://example.com/\u0085path',
		'https://example.com/\u202epath',
	])('rejects raw control characters without deleting the diagnostic original', (address) => {
		const input = sheet(link(cell('Address', 'placeholder')));
		input.sections.get('Hyperlink:0')!.rows.get('Link')!.cells.set('Address', { value: address });
		expect(shapeMetadata(input, () => {}).hyperlinks[0]).toMatchObject({
			address,
			target: { kind: 'unresolved' },
		});
	});
	it('keeps internal page/shape subaddresses unresolved as identifiers and never fetches', () => {
		const fetch = vi.fn(() => {
			throw new Error('Must not fetch');
		});
		vi.stubGlobal('fetch', fetch);
		try {
			const result = read(link(cell('Address', '') + cell('SubAddress', 'Page-2/Sheet.42')));
			expect(result.hyperlinks[0]!.target).toEqual({
				kind: 'internal',
				subAddress: 'Page-2/Sheet.42',
			});
			expect(result.hyperlinks[0]).not.toHaveProperty('pageId');
			expect(fetch).not.toHaveBeenCalled();
		} finally {
			vi.unstubAllGlobals();
		}
	});
	it('rejects stale/formula-only addresses and unsupported resolution options', () => {
		for (const contents of [
			'<Cell N="Address" F="NOW()"/>',
			'<Cell N="Address" V="https://example.com" E="#REF!"/>',
		]) {
			expect(read(link(contents + cell('SubAddress', 'Page-2'))).hyperlinks[0]!.target).toEqual({
				kind: 'unresolved',
				reason: 'missing-or-stale-cache',
			});
		}
		expect(
			read(link(cell('Address', 'https://example.com') + cell('ExtraInfo', '?x=1'))).hyperlinks[0]!
				.target,
		).toEqual({ kind: 'unresolved', reason: 'unsupported-options' });
	});
	it.each(['https://example.com/guide', 'mailto:person@example.com'])(
		'retains external subaddresses without exposing an incomplete link for %s',
		(address) => {
			const result = read(link(cell('Address', address) + cell('SubAddress', 'install')));
			expect(result.hyperlinks[0]).toMatchObject({
				address,
				subAddress: 'install',
				target: { kind: 'unresolved', reason: 'unsupported-external-subaddress' },
			});
			expect(result.reports).toContain('unresolved-hyperlink');
		},
	);
	it('retains multiple rows in order, without picking one of several default links', () => {
		const result = read(
			section(
				'Hyperlink',
				named('First', cell('SubAddress', 'One') + cell('Default', 1)) +
					named('Second', cell('SubAddress', 'Two') + cell('Default', 1)),
			),
		);
		expect(result.hyperlinks.map((item) => [item.name, item.default])).toEqual([
			['First', true],
			['Second', true],
		]);
	});
});

describe('shape metadata resource limits and parser integration', () => {
	it('rejects strings exceeding 8192 UTF-16 code units and accepts the boundary', () => {
		expect(() => read(prop(named('Text', cell('Value', 'x'.repeat(8192)))))).not.toThrow();
		expect(() => read(prop(named('Text', cell('Value', 'x'.repeat(8193)))))).toThrow(
			expect.objectContaining({ code: 'METADATA_LIMIT' }),
		);
	});
	it('counts normalized URL output and repeated raw/display text in the character budget', () => {
		const input = sheet(prop(named('x', cell('Value', 'a'))));
		expect(() =>
			shapeMetadata(input, () => {}, createMetadataBudget({ maxCharacters: 4 })),
		).not.toThrow();
		expect(() =>
			shapeMetadata(input, () => {}, createMetadataBudget({ maxCharacters: 3 })),
		).toThrow(expect.objectContaining({ code: 'METADATA_LIMIT' }));
		const external = sheet(link(cell('Address', 'https://例子.测试/路径')));
		expect(() =>
			shapeMetadata(external, () => {}, createMetadataBudget({ maxCharacters: 50 })),
		).toThrow(expect.objectContaining({ code: 'METADATA_LIMIT' }));
	});
	it('bounds combined per-shape rows and aggregate rows shared across inherited instances', () => {
		const input = sheet(prop(named('x', cell('Value', 'a')))),
			budget = createMetadataBudget({ maxRows: 1 });
		shapeMetadata(input, () => {}, budget);
		expect(() => shapeMetadata(input, () => {}, budget)).toThrow(
			expect.objectContaining({ code: 'METADATA_LIMIT' }),
		);
		expect(() =>
			read(
				prop(Array.from({ length: 1025 }, (_, i) => named(`P${i}`, cell('Value', 'x'))).join('')),
			),
		).toThrow(expect.objectContaining({ code: 'METADATA_LIMIT' }));
	});
	it.each([0, -1, 1.5, Infinity, NaN, Number.MAX_SAFE_INTEGER + 1])(
		'rejects invalid budget option %s',
		(limit) => {
			expect(() => createMetadataBudget({ maxRows: limit })).toThrow(
				expect.objectContaining({ code: 'INVALID_LIMIT' }),
			);
			expect(() => createMetadataBudget({ maxCharacters: limit })).toThrow(
				expect.objectContaining({ code: 'INVALID_LIMIT' }),
			);
		},
	);
	it('parses master data overrides and inherited hyperlinks, without inheriting formatting style data', async () => {
		const data = prop(
			named('Asset', cell('Type', 2) + cell('Value', 10) + cell('Label', 'Master label')),
		);
		const bytes = await fixture({
			document: `<StyleSheets><StyleSheet ID="4">${prop(named('StyleOnly', cell('Value', 'no')))}</StyleSheet></StyleSheets>`,
			masters: [
				{ id: '9', shapes: shape('99', data + link(cell('Address', 'https://example.com'))) },
			],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', prop(named('Asset', cell('Value', 42))), 'Master="9" TextStyle="4"')}</Shapes>`,
				},
			],
		});
		const item = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		expect(item.shapeData).toHaveLength(1);
		expect(item.shapeData![0]).toMatchObject({
			name: 'Asset',
			label: 'Master label',
			type: 2,
			value: 42,
		});
		expect(item.hyperlinks![0]).toMatchObject({
			target: { kind: 'external', href: 'https://example.com/' },
		});
	});
	it('enforces document-wide character limits across repeated master instances', async () => {
		const bytes = await fixture({
			masters: [{ id: '9', shapes: shape('99', prop(named('x', cell('Value', 'a')))) }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', '', 'Master="9"')}${shape('2', '', 'Master="9"')}</Shapes>`,
				},
			],
		});
		await expect(parseVsdx(bytes, { metadata: { maxCharacters: 7 } })).rejects.toMatchObject({
			code: 'METADATA_LIMIT',
		});
		await expect(parseVsdx(bytes, { metadata: { maxCharacters: 8 } })).resolves.toHaveProperty(
			'format',
			'vsdx',
		);
	});
	it('enforces document-wide row limits across repeated master instances on separate pages', async () => {
		const bytes = await fixture({
			masters: [{ id: '9', shapes: shape('99', prop(named('Data', cell('Value', 'one')))) }],
			pages: [0, 1].map((i) => ({
				id: String(i),
				contents: `<Shapes>${shape(String(i + 1), '', 'Master="9"')}</Shapes>`,
			})),
		});
		await expect(parseVsdx(bytes, { metadata: { maxRows: 1 } })).rejects.toMatchObject({
			code: 'METADATA_LIMIT',
		});
	});
});
