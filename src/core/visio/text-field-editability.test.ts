import { expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { fixture, shape } from './test-fixtures';
const row = '<Section N="Field"><Row IX="0"><Cell N="Value" V="2" F="Width"/></Row></Section>';
const text = '<Text><fld IX="0">2</fld></Text>';
it.each([
	'plain',
	'master',
	'group',
	'child',
	'duplicate',
	'guard',
	'calendar',
	'oversized',
	'locked',
])('projects source-backed edit eligibility for %s', async (kind) => {
	let contents = row + text;
	if (kind === 'duplicate') contents = row + '<Text><fld IX="0">2</fld><fld IX="0">2</fld></Text>';
	if (kind === 'guard') contents = contents.replace('F="Width"', 'F="GUARD(Width)"');
	if (kind === 'calendar')
		contents = contents.replace('</Row>', '<Cell N="Calendar" V="1"/></Row>');
	if (kind === 'oversized')
		contents = contents.replace('</Row>', `<Cell N="Unknown" V="${'x'.repeat(3000)}"/></Row>`);
	if (kind === 'locked') contents += '<Cell N="LockTextEdit" V="1"/>';
	if (kind === 'group') contents += '<Shapes/>';
	const local = kind === 'master' ? shape('1', '', 'Master="1"') : shape('1', contents);
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes>' +
					(kind === 'child'
						? shape('2', '<Shapes>' + local + '</Shapes>', 'Type="Group"')
						: local) +
					'</Shapes>',
			},
		],
		...(kind === 'master' ? { masters: [{ id: '1', shapes: shape('9', row + text) }] } : {}),
	});
	const page = (await parseVsdx(bytes)).pages[0]!;
	const parsed = kind === 'child' ? page.shapes[0]!.children[0]! : page.shapes[0]!;
	expect(parsed.text.fields?.length).toBeGreaterThan(0);
	for (const field of parsed.text.fields!) expect(field.editable === true).toBe(kind === 'plain');
});
