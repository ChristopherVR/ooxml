// Windows acceptance inputs. Run from src/core: bun scripts/shared-picture-native-fixtures.ts <output-dir>
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createDocument, saveDocx } from '../docx/index';
import { loadXlsx } from '../xlsx/read/index';
import { saveXlsx } from '../xlsx/write/index';
import { createEditSession } from '../xlsx/edit/session';
const directory = resolve(process.argv[2] ?? '.native-picture-acceptance');
mkdirSync(directory, { recursive: true });
const model = createDocument();
// The local path is twice the shape's width. A shape-based guide at w/4 appears at 1/8 display width.
const customGeometryXml =
	'<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:avLst/><a:gdLst><a:gd name="x" fmla="*/ w 1 4"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path w="3810000" h="762000"><a:moveTo><a:pt x="x" y="0"/></a:moveTo><a:lnTo><a:pt x="3810000" y="762000"/></a:lnTo><a:lnTo><a:pt x="0" y="762000"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
model.blocks = [
	{
		type: 'paragraph',
		id: 'p',
		runs: [
			{
				text: '',
				image: {
					relId: '',
					partName: '',
					contentType: 'application/octet-stream',
					widthPx: 200,
					heightPx: 80,
					anchored: true,
					unsupported: 'Shape',
					shape: { preset: 'custom', customGeometryXml, editable: true },
				},
			},
		],
	},
];
writeFileSync(resolve(directory, 'custom-shape.docx'), await saveDocx(model));
const workbook = await loadXlsx(
	readFileSync(new URL('../xlsx/__fixtures__/excel-features.xlsx', import.meta.url)),
);
const sheet = workbook.sheets.findIndex((value) =>
	value.drawings.some((drawing) => drawing.kind === 'image'),
);
const index = workbook.sheets[sheet]!.drawings.findIndex((drawing) => drawing.kind === 'image');
const session = createEditSession(workbook);
session.setDrawingAnchor(sheet, index, {
	from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
	ext: { cx: 952500, cy: 476250 },
});
session.setPictureCrop(sheet, index, { cropLeft: 0.25, cropRight: 0.1 });
session.editPictureEffects(sheet, index, {
	grayscale: true,
	brightness: 0.2,
	contrast: -0.1,
	alpha: 0.5,
});
writeFileSync(resolve(directory, 'picture-effects.xlsx'), await saveXlsx(workbook));
console.log(directory);
