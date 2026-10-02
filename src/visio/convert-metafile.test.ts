import { describe, expect, it, vi } from 'vitest';
import { convertVisioMetafile, type VisioMetafileTreeConverter } from './convert-metafile.js';
import { emf, record, setWord } from './emf-admission.test-fixtures.js';
const tree = () => ({
	tag: 'svg',
	attrs: { xmlns: 'http://www.w3.org/2000/svg', width: 100, height: 100, viewBox: '0 0 100 100' },
	children: [],
});
describe('bounded package conversion adapter', () => {
	it('copies only the selected input view and passes fixed safe converter options', async () => {
		const source = emf([record(43, [0, 0, 20, 20])]);
		const padded = new Uint8Array(source.length + 8);
		padded.set(source, 4);
		const convert = vi.fn<VisioMetafileTreeConverter>(async (buffer: ArrayBuffer) => {
			expect(new Uint8Array(buffer)).toEqual(source);
			expect(buffer).not.toBe(padded.buffer);
			return tree();
		});
		const result = await convertVisioMetafile(padded.subarray(4, padded.length - 4), convert);
		expect(result.status).toBe('ok');
		expect(convert.mock.calls[0]?.[1]).toMatchObject({
			exactRasterOps: false,
			dpiScale: 1,
			maxCanvasDimension: 2048,
		});
	});
	it.each([
		new Uint8Array(),
		emf([record(0xffff)]),
		emf([record(33), record(34, [-1])]),
		new Uint8Array(256 * 1024 + 1),
	])('never calls the package for rejected input', async (bytes) => {
		const convert = vi.fn();
		expect((await convertVisioMetafile(bytes, convert)).status).not.toBe('ok');
		expect(convert).not.toHaveBeenCalled();
	});
	it('rejects shared input and never invokes caller accessors', async () => {
		const bytes = emf([]),
			getter = vi.fn(() => {
				throw new Error('getter');
			});
		Object.defineProperty(bytes, 'byteLength', { get: getter });
		expect((await convertVisioMetafile(bytes, async () => tree())).status).toBe('ok');
		expect(getter).not.toHaveBeenCalled();
		expect(
			(await convertVisioMetafile(new Uint8Array(new SharedArrayBuffer(128)), vi.fn())).status,
		).toBe('invalid');
	});
	it.each([
		null,
		{ ...tree(), attrs: { ...tree().attrs, width: 99, viewBox: '0 0 99 100' } },
		{ ...tree(), children: [{ tag: 'script', attrs: {}, children: [] }] },
	])('rejects null, wrong dimensions and active output', async (output) => {
		expect((await convertVisioMetafile(emf([]), async () => output)).status).toBe(
			'conversion-failed',
		);
	});
	it.each([setWord(emf([]), 16, 2048), emf(Array.from({ length: 511 }, () => record(27, [0, 0])))])(
		'rejects excessive dimensions or records before package execution',
		async (bytes) => {
			const convert = vi.fn();
			expect((await convertVisioMetafile(bytes, convert)).status).toBe('budget-exceeded');
			expect(convert).not.toHaveBeenCalled();
		},
	);
	it('reports oversized output as budget exceeded', async () => {
		const output = { ...tree(), children: new Array(25_001).fill(null) };
		expect((await convertVisioMetafile(emf([]), async () => output)).status).toBe(
			'budget-exceeded',
		);
	});
	it('redacts converter errors', async () => {
		const result = await convertVisioMetafile(emf([]), async () => {
			throw new Error('private document contents');
		});
		expect(result.status).toBe('conversion-failed');
		expect(JSON.stringify(result)).not.toContain('private document');
	});
	it('isolates input changes during async package work', async () => {
		const source = emf([]);
		await convertVisioMetafile(source, async (buffer) => {
			const before = new Uint8Array(buffer).slice();
			source.fill(0);
			await Promise.resolve();
			expect(new Uint8Array(buffer)).toEqual(before);
			return tree();
		});
	});
});

it('applies the tighter conversion output ceiling before retaining a document asset', async () => {
	const output = {
		...tree(),
		children: Array.from({ length: 2048 }, () => ({ tag: 'g', attrs: {}, children: [] })),
	};
	const result = await convertVisioMetafile(emf([]), async () => output);
	expect(result).toMatchObject({ status: 'budget-exceeded', code: 'vector-limit' });
});
