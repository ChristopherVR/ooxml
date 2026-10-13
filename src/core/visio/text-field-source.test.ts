import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { fixture, shape } from './test-fixtures';
import { readCells } from './sheet';
import { parseXml } from '../xml';
import { textFieldSourceKey } from './text-field-source';
it('shares deterministic bounded source proofs for repeated immutable row maps', () => {
	const cells = readCells(
		parseXml(
			'<Row xmlns="http://schemas.microsoft.com/office/visio/2012/main"><Cell N="Value" V="1"/><Cell N="Type" V="2"/></Row>',
		).documentElement,
	);
	expect(textFieldSourceKey(cells)).toBe(
		'[["Type","2",null,null,null],["Value","1",null,null,null]]',
	);
	expect(textFieldSourceKey(cells)).toBe(textFieldSourceKey(cells));
});
it.each(['large', 'many'])(
	'retains display without edit proofs for %s Field rows',
	async (kind) => {
		const extra =
			kind === 'large'
				? `<Cell N="Unknown" V="${'x'.repeat(3000)}"/>`
				: Array.from({ length: 65 }, (_, i) => `<Cell N="Unknown${i}" V="0"/>`).join('');
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents:
						'<Shapes>' +
						shape(
							'1',
							'<Section N="Field"><Row IX="0"><Cell N="Value" V="2" F="2"/>' +
								extra +
								'</Row></Section><Text><fld IX="0">2</fld><fld IX="0">2</fld></Text>',
						) +
						'</Shapes>',
				},
			],
		});
		const text = (await parseVsdx(bytes)).pages[0]!.shapes[0]!.text;
		expect(text.plainText).toBe('22');
		expect(text.fields).toHaveLength(2);
		for (const field of text.fields!) expect(field.sourceKey).toBeUndefined();
	},
);
