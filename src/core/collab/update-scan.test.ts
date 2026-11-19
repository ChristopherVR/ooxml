import * as encoding from 'lib0/encoding';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';

import { applyUpdateSafe, validateUpdate } from './codec';
import { MIN_BYTES_PER_STRUCT, scanUpdate } from './update-scan';

/** A document using every kind of content Yjs encodes. */
function richDoc(): Y.Doc {
	const doc = new Y.Doc();
	const map = doc.getMap('m');
	map.set('number', 42);
	map.set('negative', -7);
	map.set('float', 1.5);
	map.set('big', 2n ** 40n);
	map.set('flags', [true, false, null, undefined]);
	map.set('nested', { a: { b: { c: [1, 'x', new Uint8Array([1, 2])] } } });
	map.set('bytes', new Uint8Array([9, 8, 7]));
	const inner = new Y.Map();
	map.set('inner', inner);
	inner.set('k', 'v');
	const text = doc.getText('t');
	text.insert(0, 'hello world');
	text.format(0, 5, { bold: true });
	text.insertEmbed(3, { image: 'x' });
	text.delete(6, 2);
	const array = doc.getArray('a');
	array.insert(0, [1, 2, 3, 'four', { five: 5 }]);
	array.delete(1, 1);
	const xml = doc.getXmlFragment('x');
	const element = new Y.XmlElement('p');
	element.setAttribute('class', 'c');
	element.insert(0, [new Y.XmlText('inside')]);
	xml.insert(0, [element]);
	map.set('hook', new Y.XmlHook('hook'));
	array.push([new Y.Doc({ guid: 'sub' })]);
	map.delete('number');
	return doc;
}

/** The structure of the audit's frame: one client declaring `count` two-byte GC structs. */
function gcBomb(count: number): Uint8Array {
	const e = encoding.createEncoder();
	encoding.writeVarUint(e, 1);
	encoding.writeVarUint(e, count);
	encoding.writeVarUint(e, 987654);
	encoding.writeVarUint(e, 0);
	for (let i = 0; i < count; i++) {
		encoding.writeUint8(e, 0);
		encoding.writeVarUint(e, 1);
	}
	encoding.writeVarUint(e, 0);
	return encoding.toUint8Array(e);
}

describe('scanUpdate', () => {
	it('reads every kind of content exactly as Yjs encodes it', () => {
		const doc = richDoc();
		for (const update of [Y.encodeStateAsUpdate(doc), Y.encodeStateAsUpdate(new Y.Doc())]) {
			const scan = scanUpdate(update);
			expect(scan).toMatchObject({ ok: true, bytesRead: update.byteLength });
		}
		const updates: Uint8Array[] = [];
		const live = new Y.Doc();
		live.on('update', (update: Uint8Array) => updates.push(update));
		live.getText('t').insert(0, 'abc');
		live.getText('t').delete(1, 1);
		live.getMap('m').set('k', { deep: [1, 2] });
		for (const update of updates) {
			expect(scanUpdate(update)).toMatchObject({ ok: true, bytesRead: update.byteLength });
		}
	});

	it('refuses a small frame that would allocate millions of structs, before decoding it', () => {
		const bomb = gcBomb(3_000_000);
		expect(bomb.byteLength).toBeLessThan(7 * 1024 * 1024);
		expect(scanUpdate(bomb)).toEqual({
			ok: false,
			reason: `update holds more than ${Math.ceil(bomb.byteLength / MIN_BYTES_PER_STRUCT)} structs`,
			overLimit: true,
		});
		expect(validateUpdate(bomb).ok).toBe(false);
		const doc = new Y.Doc();
		expect(applyUpdateSafe(doc, bomb).ok).toBe(false);
		expect(doc.store.clients.size).toBe(0);
	});

	it('counts values inside JSON and any contents and delete ranges', () => {
		const doc = new Y.Doc();
		doc.getArray('a').insert(
			0,
			Array.from({ length: 50 }, (_, i) => ({ i })),
		);
		const update = Y.encodeStateAsUpdate(doc);
		expect(scanUpdate(update, { maxStructs: 20 }).ok).toBe(false);
		const text = new Y.Doc();
		const t = text.getText('t');
		t.insert(0, 'x'.repeat(200));
		for (let i = 0; i < 50; i++) t.delete(i * 2, 1);
		expect(scanUpdate(Y.encodeStateAsUpdate(text), { maxDeleteRanges: 10 }).ok).toBe(false);
	});

	it('admits an update that only deletes characters spread through a short text', () => {
		// What replacing text in a deck sends: many one-character ranges at small clocks, about two
		// bytes each, refused before deleted ranges had their own budget.
		const doc = new Y.Doc();
		const text = doc.getText('t');
		for (let i = 0; i < 40; i++) text.insert(i, 'x');
		const updates: Uint8Array[] = [];
		doc.on('update', (update: Uint8Array) => updates.push(update));
		doc.transact(() => {
			for (let i = 39; i >= 0; i -= 2) text.delete(i, 1);
		});
		const [update] = updates as [Uint8Array];
		const scan = scanUpdate(update);
		expect(scan).toMatchObject({ ok: true, deleteRanges: 21 });
		expect(update.byteLength).toBeLessThan(21 * 4);
	});

	it('admits long lists of primitive values', () => {
		// Primitive values cost no budget: a long list of small numbers is about one byte each.
		const numbers = new Y.Doc();
		numbers.getArray('a').insert(
			0,
			Array.from({ length: 100_000 }, (_, i) => i % 50),
		);
		expect(scanUpdate(Y.encodeStateAsUpdate(numbers)).ok).toBe(true);
	});

	it('refuses truncated or malformed encodings', () => {
		const update = Y.encodeStateAsUpdate(richDoc());
		for (const cut of [1, 3, Math.floor(update.byteLength / 2), update.byteLength - 2]) {
			const scan = scanUpdate(update.subarray(0, cut));
			if (scan.ok) expect(scan.bytesRead).toBeLessThanOrEqual(cut);
		}
		expect(scanUpdate(new Uint8Array([1, 1, 1, 0, 0x1f]))).toMatchObject({
			ok: false,
			overLimit: false,
		});
		expect(
			scanUpdate(new Uint8Array([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff])),
		).toMatchObject({ ok: false });
	});
});
