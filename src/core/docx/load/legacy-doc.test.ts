import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import { LEGACY_DOC_NOTICE, LegacyDocError, loadLegacyDoc } from './legacy-doc';
import { buildOle2 } from '@christophervr/ole2/ole2-parser-write';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { ENTRY_TYPE_ROOT, ENTRY_TYPE_STREAM } from '@christophervr/ole2/ole2-parser-types';
import { readCompoundFileStream } from '@christophervr/ole2/ole2-stream-edit';
import { readDocFib } from '@christophervr/ole2/ole-document-doc-fib';

const fixture = new URL('./__fixtures__/ole-word-97.doc', import.meta.url);

// Authored CFB ownership regression: nested duplicate occurs before root stream
// in the directory, while each storage has a separate valid sibling tree.
function nestedWordFixture(source: Uint8Array, rootEncrypted: boolean): Uint8Array {
	const parsed = parseOle2(
		source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer,
	);
	const rootWord = parsed.getStream('WordDocument')!.slice();
	const tableName = readDocFib(rootWord).tableStreamName;
	const nestedWord = rootWord.slice();
	for (const [bytes, encrypted] of [
		[rootWord, rootEncrypted],
		[nestedWord, !rootEncrypted],
	] as const) {
		const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
		view.setUint16(0x0a, (view.getUint16(0x0a, true) & ~0x8100) | (encrypted ? 0x100 : 0), true);
	}
	const out = new Uint8Array(
		buildOle2(
			new Map([
				['NestedWord', nestedWord],
				['ObjectPool', new Uint8Array()],
				[tableName, parsed.getStream(tableName)!],
				['WordDocument', rootWord],
			]),
		),
	);
	const view = new DataView(out.buffer);
	const directory = view.getUint32(0x30, true),
		fat = view.getUint32(0x4c, true);
	// This tiny builder output has two consecutive directory sectors.
	expect(view.getUint32((fat + 1) * 512 + directory * 4, true)).toBe(directory + 1);
	const entry = (id: number) => (directory + 1) * 512 + id * 128;
	const built = parseOle2(out.buffer);
	const id = (name: string) => built.entries.findIndex((entry) => entry.name === name);
	const nestedId = id('NestedWord'),
		storageId = id('ObjectPool'),
		tableId = id(tableName),
		rootId = id('WordDocument');
	expect(nestedId).toBeLessThan(rootId);
	const links = (id: number, left: number, right: number, child = 0xffffffff) => {
		view.setUint32(entry(id) + 68, left, true);
		view.setUint32(entry(id) + 72, right, true);
		view.setUint32(entry(id) + 76, child, true);
		out[entry(id) + 67] = 1; // black
	};
	for (let index = 0; index < 32; index++) view.setUint16(entry(nestedId) + index * 2, 0, true);
	for (const [index, character] of Array.from('WordDocument').entries())
		view.setUint16(entry(nestedId) + index * 2, character.charCodeAt(0), true);
	view.setUint16(entry(nestedId) + 64, 26, true);
	out[entry(storageId) + 66] = 1; // ObjectPool storage
	links(0, 0xffffffff, 0xffffffff, storageId);
	links(nestedId, 0xffffffff, 0xffffffff);
	links(storageId, tableId, rootId, nestedId);
	links(tableId, 0xffffffff, 0xffffffff);
	links(rootId, 0xffffffff, 0xffffffff);
	return out;
}

describe('legacy .doc', () => {
	it.each([false, true])(
		'classifies only root WordDocument encryption (%s), not an earlier nested alias',
		async (encrypted) => {
			const bytes = nestedWordFixture(new Uint8Array(await readFile(fixture)), encrypted);
			const root = readCompoundFileStream(bytes, ['WordDocument']);
			expect(root).toBeDefined();
			if (!encrypted) {
				const fib = readDocFib(root!);
				expect(readCompoundFileStream(bytes, [fib.tableStreamName])).toBeDefined();
			}
			if (encrypted) await expect(loadLegacyDoc(bytes)).rejects.toThrow(/Encrypted legacy Word/);
			else expect((await loadLegacyDoc(bytes)).model.blocks.length).toBeGreaterThan(0);
		},
	);
	it('opens a Word 97 document as a docx model and saves it as a docx package', async () => {
		const bytes = new Uint8Array(await readFile(fixture));
		const loaded = await loadLegacyDoc(bytes);
		expect(loaded.model.blocks.length).toBeGreaterThan(0);
		expect(loaded.model.blocks[0]?.type).toBe('paragraph');
		expect(loaded.model.warnings[0]).toBe(LEGACY_DOC_NOTICE);
		const saved = await loaded.save();
		expect(Array.from(saved.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
		const reopened = await loadDocx(saved);
		expect(reopened.model.blocks.length).toBe(loaded.model.blocks.length);
	});

	it('rejects truncated and non-OLE inputs clearly', async () => {
		await expect(loadLegacyDoc(new Uint8Array([0xd0, 0xcf]))).rejects.toThrow(
			/Invalid or truncated/,
		);
		await expect(loadLegacyDoc(new Uint8Array(512))).rejects.toThrow(LegacyDocError);
	});

	it.each([1 << 8, 1 << 15])(
		'reports encrypted or obfuscated legacy Word flag %i explicitly',
		async (encryptionFlag) => {
			const source = new Uint8Array(await readFile(fixture));
			const parsed = parseOle2(
				source.buffer.slice(
					source.byteOffset,
					source.byteOffset + source.byteLength,
				) as ArrayBuffer,
			);
			const streams = new Map<string, Uint8Array>();
			for (const entry of parsed.entries) {
				if (entry.type === ENTRY_TYPE_STREAM) {
					const stream = parsed.getStream(entry.name);
					if (stream) streams.set(entry.name, stream);
				}
			}
			const word = streams.get('WordDocument');
			if (!word) throw new Error('fixture has no WordDocument stream');
			const flags = new DataView(word.buffer, word.byteOffset, word.byteLength).getUint16(
				0x0a,
				true,
			);
			new DataView(word.buffer, word.byteOffset, word.byteLength).setUint16(
				0x0a,
				flags | encryptionFlag,
				true,
			);
			streams.set('WordDocument', word);
			const clsid = parsed.entries.find((entry) => entry.type === ENTRY_TYPE_ROOT)?.clsid;
			const encrypted = new Uint8Array(buildOle2(streams, clsid));
			await expect(loadLegacyDoc(encrypted)).rejects.toThrow(/Encrypted legacy Word/);
		},
	);
});
