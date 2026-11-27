import { readVisioEmbeddedObject } from '../../../visio/ole-object';
import { parseVsdx } from '../../../visio/parser';
import { fixture, relations, shape } from '../../../visio/test-fixtures';
import type { SecurityCase } from '../case';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** A page with one embedded object: `target` from the page, declared as `contentType`. */
const drawing = (
	target: string,
	contentType: string,
	bytes: Uint8Array,
	rel = 'package',
	mode = '',
) =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape(
					'1',
					'<Cell N="PinX" V="1"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/><ForeignData ForeignType="Object" ObjectType="33280"><Rel r:id="rId9"/></ForeignData>',
					'Type="Foreign"',
				)}</Shapes>`,
			},
		],
		edit: (zip) => {
			zip.file(
				'[Content_Types].xml',
				`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>${mode ? '' : `<Override PartName="/${target}" ContentType="${contentType}"/>`}</Types>`,
			);
			zip.file(
				'visio/pages/_rels/page1.xml.rels',
				relations(
					`<Relationship Id="rId9" Type="${R}/${rel}" Target="${mode ? target : `../${target.slice(6)}`}"${mode ? ` TargetMode="${mode}"` : ''}/>`,
				),
			);
			if (!mode) zip.file(target, bytes);
		},
	});

const refusal = async (run: () => Promise<unknown>) => {
	try {
		await run();
		return 'accepted';
	} catch (error) {
		return error instanceof Error && 'code' in error ? [error.code, error.message] : 'unexpected';
	}
};

export const cases: readonly SecurityCase[] = [
	{
		name: 'an embedded object that is not an Office package is never handed out to open',
		limits: { ms: 3000, heapMb: 64 },
		run: async () => {
			// An executable header behind a Word content type and a .docx name.
			const bytes = await drawing(
				'visio/embeddings/Report.docx',
				DOCX,
				new Uint8Array([0x4d, 0x5a, 0x90, 0, 3, 0, 0, 0]),
			);
			return refusal(() => readVisioEmbeddedObject(bytes, 'visio/embeddings/Report.docx'));
		},
		expected: ['EDIT_TARGET_NOT_FOUND', 'The embedded document is not a valid Office package.'],
	},
	{
		name: 'a malformed binary OLE container is kept inert and never opened',
		limits: { ms: 3000, heapMb: 64 },
		run: async () => {
			const bytes = await drawing(
				'visio/embeddings/oleObject1.bin',
				'application/vnd.openxmlformats-officedocument.oleObject',
				new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0xff, 0xff]),
				'oleObject',
			);
			const model = await parseVsdx(bytes);
			const object = model.pages[0]?.shapes[0]?.oleObject;
			return [
				object?.format,
				Boolean(object?.refusal),
				await refusal(() => readVisioEmbeddedObject(bytes, 'visio/embeddings/oleObject1.bin')),
			];
		},
		expected: [
			'ole',
			true,
			[
				'EDIT_TARGET_NOT_FOUND',
				'This object is stored for the program that created it; it is shown from its saved picture and kept as it is, but cannot be opened here.',
			],
		],
	},
	{
		name: 'a linked OLE object is never loaded',
		limits: { ms: 3000, heapMb: 64 },
		run: async () => {
			const bytes = await drawing(
				'https://example.test/a.docx',
				DOCX,
				new Uint8Array(),
				'package',
				'External',
			);
			const model = await parseVsdx(bytes);
			return [
				model.pages[0]?.shapes[0]?.oleObject ?? null,
				model.diagnostics.some((note) => note.code === 'ole-object-linked'),
			];
		},
		expected: [null, true],
	},
];
