import { expect, it } from 'vitest';
import { preserveAttributeWhitespace } from './xml-attribute-whitespace';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { fixture } from './test-fixtures';

it('escapes attribute whitespace while preserving comments, CDATA and processing instructions', () => {
	const preserved = '<!-- <tag a="a\nb"/> --><![CDATA[<tag a="a\nb"/>]]><?keep value="a\nb"?>';
	expect(preserveAttributeWhitespace(`<root a="a\nb>c\td\re">${preserved}\n</root>`)).toBe(
		`<root a="a&#10;b>c&#9;d&#13;e">${preserved}\n</root>`,
	);
});

it('keeps character-reference whitespace in unknown attributes when editing unrelated text', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents:
					'<Shapes><Shape ID="1" Extra="before&#10;after&#9;tab&#13;return"><Text>Old</Text></Shape></Shapes>',
			},
		],
	});
	const result = await editVsdx(bytes, [
		{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'New' },
	]);
	const pkg = await VisioPackage.open(result.bytes);
	const xml = new TextDecoder().decode(await pkg.readBytes('visio/pages/page1.xml'));
	expect(xml).toContain('Extra="before&#10;after&#9;tab&#13;return"');
	expect(xml).toContain('<Text>New\n</Text>');
});
