/**
 * UserEditAtom chain / persist directory tests with a synthetic
 * incremental-save stream: the newer edit's directory entries must override
 * the older edit's.
 */
import { describe, expect, it } from 'vitest';

import { buildPersistDirectory, parsePersistDirectoryAtom } from './persist-directory';
import { PptParseError } from './record-stream';
import { RT } from './record-types';

class StreamBuilder {
	private readonly parts: Uint8Array[] = [];
	private length = 0;

	public get offset(): number {
		return this.length;
	}

	public push(bytes: Uint8Array): number {
		const at = this.length;
		this.parts.push(bytes);
		this.length += bytes.length;
		return at;
	}

	public build(): DataView {
		const out = new Uint8Array(this.length);
		let offset = 0;
		for (const part of this.parts) {
			out.set(part, offset);
			offset += part.length;
		}
		return new DataView(out.buffer);
	}
}

function header(recType: number, recLen: number, recVer = 0, recInstance = 0): Uint8Array {
	const out = new Uint8Array(8);
	const view = new DataView(out.buffer);
	view.setUint16(0, (recVer & 0x0f) | ((recInstance & 0x0fff) << 4), true);
	view.setUint16(2, recType, true);
	view.setUint32(4, recLen, true);
	return out;
}

function persistDirectoryAtom(entries: Array<[number, number[]]>): Uint8Array {
	const body: number[] = [];
	for (const [persistId, offsets] of entries) {
		body.push((persistId & 0xfffff) | (offsets.length << 20));
		body.push(...offsets);
	}
	const out = new Uint8Array(8 + body.length * 4);
	out.set(header(RT.PersistDirectoryAtom, body.length * 4));
	const view = new DataView(out.buffer);
	body.forEach((value, i) => view.setUint32(8 + i * 4, value, true));
	return out;
}

function userEditAtom(offsetLastEdit: number, offsetPersistDirectory: number): Uint8Array {
	const out = new Uint8Array(8 + 0x1c);
	out.set(header(RT.UserEditAtom, 0x1c));
	const view = new DataView(out.buffer);
	view.setUint32(8, 0, true); // lastSlideIdRef
	view.setUint32(12, 0x03000000, true); // version fields
	view.setUint32(16, offsetLastEdit, true);
	view.setUint32(20, offsetPersistDirectory, true);
	view.setUint32(24, 1, true); // docPersistIdRef
	view.setUint32(28, 10, true); // persistIdSeed
	return out;
}

describe('persist-directory', () => {
	it('parses packed persist directory entries', () => {
		const stream = new StreamBuilder();
		const objects = [1, 2, 7].map(() => stream.push(header(RT.Document, 0, 0x0f)));
		const directory = stream.push(
			persistDirectoryAtom([
				[1, [objects[0]!, objects[1]!]],
				[7, [objects[2]!]],
			]),
		);
		const pairs = parsePersistDirectoryAtom(stream.build(), directory);
		expect(pairs).toStrictEqual([
			[1, objects[0]],
			[2, objects[1]],
			[7, objects[2]],
		]);
	});

	it('walks the user edit chain, newest entries winning', () => {
		const stream = new StreamBuilder();
		// Real object headers precede each save's directory and UserEditAtom.
		const document = stream.push(header(RT.Document, 0, 0x0f));
		const oldObject = stream.push(header(RT.Document, 0, 0x0f));
		const dir1 = stream.push(persistDirectoryAtom([[1, [document, oldObject]]]));
		const edit1 = stream.push(userEditAtom(0, dir1));
		const replacement = stream.push(header(RT.Document, 0, 0x0f));
		const added = stream.push(header(RT.Document, 0, 0x0f));
		const dir2 = stream.push(
			persistDirectoryAtom([
				[2, [replacement]],
				[3, [added]],
			]),
		);
		const edit2 = stream.push(userEditAtom(edit1, dir2));

		const { currentEdit, directory } = buildPersistDirectory(stream.build(), edit2);
		expect(currentEdit.docPersistIdRef).toBe(1);
		expect(directory.get(1)).toBe(document);
		expect(directory.get(2)).toBe(replacement); // newer edit overrides
		expect(directory.get(3)).toBe(added);
	});

	it('rejects circular user edit chains', () => {
		const stream = new StreamBuilder();
		const dir = stream.push(persistDirectoryAtom([[1, [0]]]));
		const editOffset = stream.offset;
		// Points at itself as "previous" edit.
		stream.push(userEditAtom(editOffset, dir));
		expect(() => buildPersistDirectory(stream.build(), editOffset)).toThrow(PptParseError);
	});

	it('rejects a wrong record type where a UserEditAtom is expected', () => {
		const stream = new StreamBuilder();
		const at = stream.push(persistDirectoryAtom([[1, [0]]]));
		expect(() => buildPersistDirectory(stream.build(), at)).toThrow(PptParseError);
	});
});
