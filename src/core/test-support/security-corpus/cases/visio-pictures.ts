import { createVsdx } from '../../../visio/create-document';
import { editVsdx } from '../../../visio/edit';
import { emf, record, setWord } from '../../../visio/emf-admission.test-fixtures';
import { bmpPicture, svgPicture, tiffPicture } from '../../../visio/picture-test-fixtures';
import type { SecurityCase } from '../case';

async function insert(bytes: Uint8Array) {
	try {
		await editVsdx(await createVsdx(), [
			{
				type: 'insert-picture',
				pageId: '0',
				shapeId: '1',
				image: bytes,
				width: 1,
				height: 1,
				x: 1,
				y: 1,
			},
		]);
		return 'accepted';
	} catch (error) {
		return error instanceof Error && 'code' in error ? [error.code, error.message] : 'unexpected';
	}
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'SVG picture insertion refuses script elements',
		limits: { ms: 2500, heapMb: 64 },
		run: () => insert(svgPicture('<script/>')),
		expected: ['INVALID_PICTURE', 'SVG picture contains active or unsupported content.'],
	},
	{
		name: 'SVG picture insertion refuses external paint references',
		limits: { ms: 2500, heapMb: 64 },
		run: () => insert(svgPicture('<rect fill="url(https://example.test/a.svg#paint)"/>')),
		expected: ['INVALID_PICTURE', 'SVG picture contains active or unsupported content.'],
	},
	{
		name: 'BMP picture insertion checks the complete pixel scanline range',
		limits: { ms: 2500, heapMb: 64 },
		run() {
			const bytes = bmpPicture().slice(0, -1);
			new DataView(bytes.buffer).setUint32(2, bytes.length, true);
			return insert(bytes);
		},
		expected: ['INVALID_PICTURE', 'Invalid or truncated BMP file.'],
	},
	{
		name: 'TIFF picture insertion checks compressed strips before decoding',
		limits: { ms: 2500, heapMb: 64 },
		run() {
			const bytes = tiffPicture(5);
			bytes.fill(0, 130);
			return insert(bytes);
		},
		expected: ['INVALID_PICTURE', 'Invalid or truncated TIFF strip.'],
	},
	{
		name: 'EMF picture insertion checks record sizes before conversion',
		limits: { ms: 2500, heapMb: 64 },
		run: () => insert(setWord(emf([record(43, [1, 1, 20, 20])]), 112, 0)),
		expected: ['INVALID_PICTURE', 'Record must progress, align, and fit entirely within input.'],
	},
];
