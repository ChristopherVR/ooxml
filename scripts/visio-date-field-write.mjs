import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseVsdx, editVsdx, VISIO_FIELD_FORMATS as F } from '../src/core/dist/visio/index.mjs';
const [sourcePath, output] = process.argv.slice(2);
if (!sourcePath || !output) throw new Error('Usage: source.vsdx output-directory');
const bytes = readFileSync(sourcePath);
const page = (await parseVsdx(bytes)).pages[0];
const cases = [];
for (const [name, format] of [
	['short', F.shortDate],
	['iso', '{{yyyy-MM-dd}}'],
	['long', F.longDate],
	['time', F.time],
	['date-time', F.dateTime],
]) {
	const saved = await editVsdx(bytes, [
		{
			type: 'insert-text-field',
			pageId: page.id,
			shapeId: page.shapes[0].id,
			formula: 'DOCCREATION()',
			format,
		},
	]);
	const path = join(output, name + '.vsdx');
	writeFileSync(path, saved.bytes);
	cases.push({
		name,
		path,
		format,
		text: (await parseVsdx(saved.bytes)).pages[0].shapes[0].text.plainText,
	});
}
writeFileSync(
	join(output, 'cases.json'),
	JSON.stringify({ timezone: Intl.DateTimeFormat().resolvedOptions().timeZone, cases }, null, 2),
);
