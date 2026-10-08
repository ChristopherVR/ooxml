import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, section, shape } from './test-fixtures';

const command = { type: 'format-text' as const, pageId: '0', shapeId: '1' };
const row = (index: number, value: string) => `<Row IX="${index}">${value}</Row>`;
const inherited = (name: string, attributes: string) => `<Cell N="${name}" F="Inh" ${attributes}/>`;
async function source(
	local: string,
	ancestor = cell('Style', 2),
	extra = '',
	parent = 'TextStyle="2"',
) {
	return fixture({
		document:
			'<StyleSheets><StyleSheet ID="2">' +
			section('Character', row(0, ancestor)) +
			section('Paragraph', row(0, cell('HorzAlign', 2))) +
			'</StyleSheet></StyleSheets>',
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					cell('Width', 4) +
						cell('Height', 2) +
						rectangle +
						extra +
						section('Character', row(0, local) + row(1, cell('Size', 0.25))) +
						section('Paragraph', row(0, inherited('HorzAlign', 'V="2"'))) +
						'<Text><pp IX="0"/><cp IX="0"/>First<cp IX="1"/>Second</Text>',
					parent,
				)}</Shapes>`,
			},
		],
		edit: (zip) => zip.file('unknown/data.bin', new Uint8Array([1, 2, 255])),
	});
}
const pageXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');

describe('proven local text formatting delegation', () => {
	it('records an explicit local choice even when it equals the inherited cache', async () => {
		const bytes = await source(inherited('Style', 'V="2"'));
		const saved = await editVsdx(bytes, [{ ...command, italic: true }]);
		expect(await pageXml(saved.bytes)).toContain('<Cell N="Style" V="2"/>');
		expect(await pageXml(saved.bytes)).not.toContain('N="Style" F="Inh"');
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
	});
	it.each(['V="2"', 'V="2.0"', ''])(
		'overrides proven local Inh cells (%s) without flattening text',
		async (attributes) => {
			const bytes = await source(inherited('Style', attributes));
			const saved = await editVsdx(bytes, [{ ...command, bold: true, horizontalAlign: 'left' }]);
			const model = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
			expect(model.text.runs.map((run) => [run.bold, run.italic])).toEqual([
				[true, true],
				[true, true],
			]);
			expect(model.text.paragraphs?.[0]?.horizontalAlign).toBe('left');
			expect(await pageXml(saved.bytes)).toContain(
				'<Text><pp IX="0"/><cp IX="0"/>First<cp IX="1"/>Second</Text>',
			);
			expect(await pageXml(saved.bytes)).not.toContain('F="Inh"');
			const original = await JSZip.loadAsync(bytes),
				after = await JSZip.loadAsync(saved.bytes);
			for (const [path, part] of Object.entries(original.files)) {
				if (part.dir || saved.changedParts.includes(path)) continue;
				expect(await after.file(path)!.async('uint8array'), path).toEqual(
					await part.async('uint8array'),
				);
			}
			expect(
				(await editVsdx(saved.bytes, [{ ...command, bold: true, horizontalAlign: 'left' }])).bytes,
			).toEqual(saved.bytes);
		},
	);
	it.each([
		['stale cache', 'V="4"', cell('Style', 2)],
		['local error', 'V="2" E="#REF!"', cell('Style', 2)],
		['ancestor error', 'V="2"', '<Cell N="Style" V="2" E="#REF!"/>'],
		['missing ancestor', 'V="2"', ''],
		['missing cacheless ancestor', '', ''],
		['incompatible cache unit', 'V="2" U="IN"', cell('Style', 2)],
		['incompatible cacheless unit', 'U="IN"', cell('Style', 2)],
		['guarded ancestor', 'V="2"', cell('Style', 2, 'GUARD(2)')],
		['stale ancestor formula', 'V="2"', cell('Style', 2, '3')],
		['unresolved ancestor', 'V="2"', cell('Style', 2, 'Inh')],
	])('refuses %s atomically', async (_name, attributes, ancestor) => {
		const bytes = await source(inherited('Style', attributes), ancestor);
		const original = bytes.slice();
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow();
		expect(bytes).toEqual(original);
	});
	it('does not extend delegation to top-level protection', async () => {
		const bytes = await source(
			inherited('Style', 'V="2"'),
			cell('Style', 2),
			cell('LockFormat', 0, 'Inh'),
		);
		await expect(editVsdx(bytes, [{ ...command, bold: true }])).rejects.toThrow(
			/inherited formatting/i,
		);
	});
});
