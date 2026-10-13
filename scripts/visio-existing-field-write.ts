import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseVsdx, editVsdx } from '../src/core/visio/index';
import { visioTextSource } from '../src/core/visio/ui/text-draft';
const [source, output] = process.argv.slice(2);
if (!source || !output)
	throw new Error('Usage: bun scripts/visio-existing-field-write.ts source.vsdx output-directory');
const bytes = readFileSync(source),
	page = (await parseVsdx(bytes)).pages[0]!,
	cases = [];
for (const [id, formula, format] of [
	['1', 'Width*2', '0.00'],
	['2', 'DOCCREATION()', '{{yyyy-MM-dd}}'],
	['3', 'TITLE()', '@'],
]) {
	const text = page.shapes.find((shape) => shape.id === id)!.text,
		field = text.fields![0]!;
	const saved = await editVsdx(bytes, [
		{
			type: 'update-text-field',
			pageId: page.id,
			shapeId: id!,
			fieldIndex: field.index!,
			formula: formula!,
			format: format!,
			expectedText: visioTextSource(text).source,
			expectedCached: field.cached,
			expectedSourceKey: field.sourceKey!,
		},
	]);
	const path = join(output, `edit-${id}.vsdx`);
	writeFileSync(path, saved.bytes);
	cases.push({
		id,
		path,
		formula,
		format,
		text: (await parseVsdx(saved.bytes)).pages[0]!.shapes.find((shape) => shape.id === id)!.text
			.plainText,
	});
}
writeFileSync(join(output, 'core-edits.json'), JSON.stringify({ cases }, null, 2));
