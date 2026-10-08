import 'fake-indexeddb/auto';
import { describe, it, expect } from 'vitest';
import { SuiteDocumentStore, suiteFileKind } from './document-store';

describe('suite document identity', () => {
	it('persists the same id across store instances and rejects stale writes', async () => {
		const name = `suite-${crypto.randomUUID()}`;
		const first = new SuiteDocumentStore(name),
			second = new SuiteDocumentStore(name);
		const doc = await first.create('brief.docx', new Uint8Array([1]));
		const edited = { ...doc, bytes: new Uint8Array([2]), revision: 2 };
		await first.commit([{ document: edited, expected: 1 }]);
		expect(Array.from((await second.get(doc.id)).bytes)).toEqual([2]);
		await expect(
			second.commit([{ document: { ...edited, bytes: new Uint8Array([3]) }, expected: 1 }]),
		).rejects.toThrow('changed');
		expect(Array.from((await first.get(doc.id)).bytes)).toEqual([2]);
	});
	it('rolls back the entire child/parent write if either revision is stale', async () => {
		const store = new SuiteDocumentStore(`suite-${crypto.randomUUID()}`);
		const parent = await store.create('report.docx', new Uint8Array([1]));
		const child = await store.create('chart.xlsx', new Uint8Array([2]), {
			parent: { id: parent.id, path: 'word/embeddings/chart.xlsx', revision: 1 },
		});
		await expect(
			store.commit([
				{ document: { ...child, revision: 2 }, expected: 1 },
				{ document: { ...parent, revision: 3 }, expected: 2 },
			]),
		).rejects.toThrow();
		expect((await store.get(child.id)).revision).toBe(1);
	});
	it('routes only supported native file types', () => {
		expect(suiteFileKind('Budget.XLSX')).toBe('xlsx');
		expect(() => suiteFileKind('legacy.doc')).toThrow();
	});
});

it('removes a parent with cached descendants and restores the same identities', async () => {
	const store = new SuiteDocumentStore(`remove-${crypto.randomUUID()}`);
	const parent = await store.create('parent.docx', new Uint8Array([1]));
	const child = await store.create('child.xlsx', new Uint8Array([2]), {
		parent: { id: parent.id, path: 'embedding', revision: 1 },
	});
	const other = await store.create('keep.docx', new Uint8Array([3]));
	await expect(store.remove(parent.id, 2)).rejects.toThrow('changed');
	expect((await store.list()).length).toBe(3);
	const removed = await store.remove(parent.id, 1);
	expect(removed.map((d) => d.id).sort()).toEqual([parent.id, child.id].sort());
	expect((await store.list()).map((d) => d.id)).toEqual([other.id]);
	await expect(store.get(child.id)).rejects.toThrow('not found');
	await store.commit(removed.map((document) => ({ document, expected: 0 })));
	expect((await store.get(child.id)).parent?.id).toBe(parent.id);
});
