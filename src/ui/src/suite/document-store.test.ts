import 'fake-indexeddb/auto';
import { describe, it, expect } from 'vitest';
import { SuiteDocumentStore, suiteFileKind } from './document-store';
import type { SuiteCloudSource, SuiteFileKind } from './document-store';

for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx'] as SuiteFileKind[]) {
	it(`${kind} derives pending cloud state for direct writes and clears it only with an acknowledged baseline`, async () => {
		const store = new SuiteDocumentStore(`cloud-${crypto.randomUUID()}`);
		const baseline = new Uint8Array([1]);
		const source: SuiteCloudSource = {
			file: {
				driveId: 'drive',
				itemId: 'item',
				name: `File.${kind}`,
				kind,
				eTag: 'v1',
				access: 'write',
			},
			baseline,
			pending: false,
		};
		const doc = await store.create(`File.${kind}`, baseline, { cloud: source });
		const edited = { ...doc, bytes: new Uint8Array([2]), revision: 2 };
		await store.commit([{ document: edited, expected: 1 }]);
		expect((await store.get(doc.id)).cloud?.pending).toBe(true);
		await store.commit([
			{
				document: {
					...edited,
					revision: 3,
					cloud: { ...source, file: { ...source.file, eTag: 'v2' }, baseline: edited.bytes },
				},
				expected: 2,
			},
		]);
		expect((await store.get(doc.id)).cloud?.pending).toBe(false);
	});
}

it('does not persist cloud capabilities or allow direct edits of read-only connected files', async () => {
	const store = new SuiteDocumentStore(`safe-cloud-${crypto.randomUUID()}`);
	const source = {
		file: {
			driveId: 'drive',
			itemId: 'item',
			name: 'File.docx',
			kind: 'docx',
			eTag: 'v1',
			access: 'read',
			webUrl: 'secret-link',
			accessToken: 'secret',
		},
		baseline: new Uint8Array([1]),
		pending: false,
	} as SuiteCloudSource;
	const doc = await store.create('File.docx', source.baseline, { cloud: source });
	expect(JSON.stringify((await store.get(doc.id)).cloud)).not.toContain('secret');
	await expect(
		store.commit([{ document: { ...doc, bytes: new Uint8Array([2]), revision: 2 }, expected: 1 }]),
	).rejects.toThrow('read-only');
	expect(Array.from((await store.get(doc.id)).bytes)).toEqual([1]);
});

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
