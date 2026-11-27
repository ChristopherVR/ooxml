import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { emf, record } from './emf-admission.test-fixtures';
import { classifyVisioOleEmbedding, readVisioEmbeddedObject } from './ole-object';
import { parseVsdx } from './parser';
import { fixture, relations, shape } from './test-fixtures';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const png = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const OLE = 'application/vnd.openxmlformats-officedocument.oleObject';

const frame =
	'<Cell N="PinX" V="2"/><Cell N="PinY" V="2"/><Cell N="Width" V="2"/><Cell N="Height" V="1"/>';

async function wordDocument(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

interface Options {
	part?: string;
	contentType?: string;
	relType?: string;
	mode?: string;
	bytes?: Uint8Array;
	preview?: { name: string; bytes: Uint8Array } | null;
	showAsIcon?: boolean;
}

async function drawing(options: Options = {}): Promise<Uint8Array> {
	const part = options.part ?? 'Microsoft_Word_Document.docx';
	const data = `<ForeignData ForeignType="Object" ObjectType="33280" ShowAsIcon="${options.showAsIcon ? 1 : 0}" ObjectWidth="2" ObjectHeight="1"><Rel r:id="rId9"/></ForeignData>`;
	const bytes = options.bytes ?? (await wordDocument());
	const preview =
		options.preview === undefined ? { name: 'image1.png', bytes: png } : options.preview;
	return fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', frame + data, 'Type="Foreign"')}</Shapes>` },
		],
		edit: (zip) => {
			zip.file(
				'[Content_Types].xml',
				`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="emf" ContentType="image/x-emf"/><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/><Override PartName="/visio/embeddings/${part}" ContentType="${options.contentType ?? DOCX}"/></Types>`,
			);
			zip.file(
				'visio/pages/_rels/page1.xml.rels',
				relations(
					`<Relationship Id="rId9" Type="${options.relType ?? `${R}/package`}" Target="../embeddings/${part}"${options.mode ? ` TargetMode="${options.mode}"` : ''}/>`,
				),
			);
			zip.file(`visio/embeddings/${part}`, bytes);
			if (preview) {
				zip.file(
					`visio/embeddings/_rels/${part}.rels`,
					relations(
						`<Relationship Id="rId1" Type="${R}/image" Target="../media/${preview.name}"/>`,
					),
				);
				zip.file(`visio/media/${preview.name}`, preview.bytes);
			}
		},
	});
}

const first = async (bytes: Uint8Array) => {
	const model = await parseVsdx(bytes);
	return { model, shape: model.pages[0]!.shapes[0]! };
};
const codes = (model: Awaited<ReturnType<typeof parseVsdx>>) =>
	model.diagnostics.map((note) => note.code);

describe('embedded OLE objects', () => {
	it('draws an embedded Word document from its saved bitmap preview and offers Open', async () => {
		const { model, shape: object } = await first(await drawing());
		expect(object.oleObject).toEqual({
			part: 'visio/embeddings/Microsoft_Word_Document.docx',
			format: 'docx',
			fileName: 'Microsoft_Word_Document.docx',
			showAsIcon: false,
		});
		expect(object.image).toMatchObject({ mimeType: 'image/png', width: 2, height: 1 });
		expect(codes(model)).not.toContain('unsupported-foreign-object');
	});

	it('keeps the icon view as stored and reads ShowAsIcon', async () => {
		const { shape: object } = await first(await drawing({ showAsIcon: true }));
		expect(object.oleObject?.showAsIcon).toBe(true);
		expect(object.image?.mimeType).toBe('image/png');
	});

	it('inspects an EMF preview through the metafile path without a bitmap', async () => {
		const { model, shape: object } = await first(
			await drawing({ preview: { name: 'image1.emf', bytes: emf([record(43, [0, 0, 90, 90])]) } }),
		);
		expect(object.oleObject?.format).toBe('docx');
		expect(object.image).toBeUndefined();
		// Without a converter the admitted EMF is reported, never claimed as drawn.
		expect(codes(model)).toContain('emf-rendering-disabled');
	});

	it('reports a missing or WMF preview honestly and keeps the object', async () => {
		const missing = await first(await drawing({ preview: null }));
		expect(missing.shape.oleObject?.format).toBe('docx');
		expect(codes(missing.model)).toContain('ole-object-no-preview');
		const wmf = await first(
			await drawing({ preview: { name: 'image1.wmf', bytes: new Uint8Array([1, 2, 3]) } }),
		);
		expect(codes(wmf.model)).toContain('ole-object-wmf-preview');
	});

	it('refuses binary OLE containers and macro-enabled documents with the reason', async () => {
		const bin = await first(
			await drawing({
				part: 'oleObject1.bin',
				contentType: OLE,
				relType: `${R}/oleObject`,
				bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]),
			}),
		);
		expect(bin.shape.oleObject).toMatchObject({ format: 'ole' });
		expect(bin.shape.oleObject?.refusal).toMatch(/cannot be opened here/);
		expect(bin.shape.image?.mimeType).toBe('image/png');
		const macro = await first(
			await drawing({
				part: 'Doc.docm',
				contentType: 'application/vnd.ms-word.document.macroEnabled.12',
			}),
		);
		expect(macro.shape.oleObject).toMatchObject({ format: 'macro-enabled' });
		expect(macro.shape.oleObject?.refusal).toMatch(/macros/);
	});

	it('never follows linked objects', async () => {
		const { model, shape: object } = await first(await drawing({ mode: 'External' }));
		expect(object.oleObject).toBeUndefined();
		expect(codes(model)).toContain('ole-object-linked');
	});

	it('reads an embedded document back for Open and refuses anything else', async () => {
		const bytes = await drawing();
		const opened = await readVisioEmbeddedObject(
			bytes,
			'visio/embeddings/Microsoft_Word_Document.docx',
		);
		expect(opened).toMatchObject({ fileName: 'Microsoft_Word_Document.docx', format: 'docx' });
		expect(Array.from(opened.bytes)).toEqual(Array.from(await wordDocument()));
		const bin = await drawing({
			part: 'oleObject1.bin',
			contentType: OLE,
			relType: `${R}/oleObject`,
			bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]),
		});
		await expect(readVisioEmbeddedObject(bin, 'visio/embeddings/oleObject1.bin')).rejects.toThrow(
			/cannot be opened here/,
		);
		const fake = await drawing({ bytes: new Uint8Array([1, 2, 3, 4]) });
		await expect(
			readVisioEmbeddedObject(fake, 'visio/embeddings/Microsoft_Word_Document.docx'),
		).rejects.toThrow(/not a valid Office package/);
		await expect(readVisioEmbeddedObject(bytes, 'visio/document.xml')).rejects.toThrow();
	});

	it('keeps the embedding and its preview byte for byte through an unrelated edit', async () => {
		const source = await drawing();
		const before = await JSZip.loadAsync(source);
		const result = await editVsdx(source, [
			{ type: 'move-shape', pageId: '0', shapeId: '1', x: 3, y: 3 },
		]);
		const after = await JSZip.loadAsync(result.bytes);
		for (const name of [
			'visio/embeddings/Microsoft_Word_Document.docx',
			'visio/embeddings/_rels/Microsoft_Word_Document.docx.rels',
			'visio/media/image1.png',
		])
			expect(await after.file(name)!.async('uint8array')).toEqual(
				await before.file(name)!.async('uint8array'),
			);
	});

	it('classifies by content type before the extension', () => {
		expect(classifyVisioOleEmbedding('a.docx', `${R}/package`, undefined)).toBe('docx');
		expect(classifyVisioOleEmbedding('a.docx', `${R}/package`, 'application/zip')).toBe('other');
		expect(classifyVisioOleEmbedding('a.bin', `${R}/oleObject`, 'application/x-other')).toBe('ole');
		expect(
			classifyVisioOleEmbedding(
				'x',
				`${R}/package`,
				'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			),
		).toBe('pptx');
	});
});
