import assert from 'node:assert/strict';
import test from 'node:test';
import {
	SuiteOneDrive,
	SuiteCloudStorage,
	cloudIdentity,
	cloudReadOnly,
	cloudSaveLabel,
} from './cloud-files.js';

const bytes = (text) => new TextEncoder().encode(text);
const file = (kind = 'docx') => ({
	driveId: 'drive',
	itemId: 'item',
	parentId: 'parent',
	name: `Example.${kind}`,
	kind,
	eTag: 'v1',
	access: 'write',
	webUrl: 'https://tenant.sharepoint.com/sensitive?capability=secret',
});
function store() {
	const records = new Map();
	return {
		async list() {
			return [...records.values()].map((doc) => structuredClone(doc));
		},
		async get(id) {
			return structuredClone(records.get(id));
		},
		async create(name, data, extra = {}) {
			const doc = {
				id: String(records.size + 1),
				name,
				kind: name.split('.').pop(),
				bytes: data,
				revision: 1,
				...extra,
			};
			records.set(doc.id, structuredClone(doc));
			return structuredClone(doc);
		},
		async commit(changes) {
			for (const { document, expected } of changes)
				assert.equal(records.get(document.id)?.revision, expected);
			for (const { document } of changes) records.set(document.id, structuredClone(document));
		},
	};
}
function harness(kind = 'docx', overrides = {}) {
	const library = store();
	const opened = { file: file(kind), bytes: bytes('baseline') };
	const client = {
		open: async () => opened,
		read: async () => opened,
		share: async (_file, options) => `https://example.invalid/${options.type}`,
	};
	const cloud = new SuiteOneDrive({
		store: library,
		client,
		createSession: () => ({ save: async () => ({ ...file(kind), eTag: 'v2' }) }),
		...overrides,
	});
	return { library, opened, cloud, client };
}

for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx']) {
	test(`${kind}: opens, persists stable identity, saves and reports cloud state`, async () => {
		const { cloud, library } = harness(kind);
		const opened = await cloud.open('https://1drv.ms/input-secret');
		assert.equal(opened.kind, kind);
		assert.equal(opened.cloud.pending, false);
		assert.equal(JSON.stringify(opened).includes('secret'), false);
		const pending = {
			...opened,
			bytes: bytes('edited'),
			cloud: { ...opened.cloud, pending: true },
		};
		await library.commit([{ document: pending, expected: opened.revision }]);
		assert.match(cloudSaveLabel(pending), /pending/);
		const saved = await cloud.save(pending);
		assert.equal(saved.cloud.file.itemId, 'item');
		assert.equal(saved.cloud.file.eTag, 'v2');
		assert.deepEqual(saved.cloud.baseline, bytes('edited'));
		assert.equal(cloudSaveLabel(saved), 'Saved to OneDrive');
		assert.deepEqual((await library.get(saved.id)).bytes, bytes('edited'));
	});
}

test('adopts combined saved bytes and reloads the editor after a semantic merge', async () => {
	const merge = async () => ({ status: 'merged', bytes: bytes('local and remote') });
	const adopted = [];
	const { cloud, library } = harness('docx', {
		merge,
		adopted: async (doc) => adopted.push(doc),
		createSession: (_client, _opened, settings) => {
			assert.equal(settings.merge, merge);
			return {
				save: async () => ({ ...file(), eTag: 'merged' }),
				getSnapshot: () => ({ baseBytes: bytes('local and remote') }),
			};
		},
	});
	const doc = await cloud.open('link');
	const pending = { ...doc, bytes: bytes('local'), cloud: { ...doc.cloud, pending: true } };
	await library.commit([{ document: pending, expected: doc.revision }]);
	const saved = await cloud.save(pending);
	assert.deepEqual(saved.bytes, bytes('local and remote'));
	assert.deepEqual(saved.cloud.baseline, bytes('local and remote'));
	assert.equal(saved.cloud.pending, false);
	assert.equal(saved.revision, doc.revision + 1);
	assert.equal(adopted.at(-1), saved);
});

test('reopening a pending or live-dirty file preserves local work', async () => {
	const { cloud, library, opened } = harness();
	const original = await cloud.open('link');
	const pending = {
		...original,
		bytes: bytes('local'),
		cloud: { ...original.cloud, pending: true },
	};
	await library.commit([{ document: pending, expected: original.revision }]);
	opened.bytes = bytes('remote');
	assert.deepEqual((await cloud.open('link')).bytes, bytes('local'));
	pending.cloud.pending = false;
	await library.commit([{ document: pending, expected: pending.revision }]);
	cloud.dirty = () => true;
	assert.deepEqual((await cloud.open('link')).bytes, bytes('local'));
});

test('failed cloud save leaves durable pending bytes for retry after reload', async () => {
	const { cloud, library } = harness('xlsx', {
		createSession: () => ({
			save: async () => {
				throw new Error('private token details');
			},
		}),
	});
	const doc = await cloud.open('link');
	doc.cloud.pending = true;
	doc.bytes = bytes('unsynchronized edits');
	await library.commit([{ document: doc, expected: doc.revision }]);
	await assert.rejects(
		cloud.save(doc),
		(error) => !error.message.includes('private') && /device/.test(error.message),
	);
	const restored = await library.get(doc.id);
	assert.equal(restored.cloud.pending, true);
	assert.deepEqual(restored.bytes, bytes('unsynchronized edits'));
	assert.equal(cloudReadOnly({ cloud: { file: { access: 'unknown' } } }), true);
});

for (const action of ['copy', 'latest', 'cancel']) {
	test(`conflict recovery ${action} retains edited bytes and never overwrites the original`, async () => {
		let adopted;
		let copies = 0;
		const remote = { file: { ...file(), eTag: 'remote' }, bytes: bytes('remote') };
		const { cloud, library } = harness('docx', {
			conflict: async () => action,
			adopted: async (doc) => {
				adopted = doc;
			},
			createSession: () => ({
				save: async () => {
					throw Object.assign(new Error('conflict'), { conflict: { remote } });
				},
				saveCopy: async (local) => {
					copies++;
					return { file: { ...file(), itemId: 'copy' }, bytes: local };
				},
				reload: async () => remote,
			}),
		});
		const doc = await cloud.open('link');
		doc.cloud.pending = true;
		doc.bytes = bytes('local');
		await library.commit([{ document: doc, expected: doc.revision }]);
		if (action === 'cancel') {
			await assert.rejects(cloud.save(doc), /newer version/);
			assert.deepEqual((await library.get(doc.id)).bytes, bytes('local'));
			assert.equal(copies, 0);
		} else {
			const saved = await cloud.save(doc);
			assert.equal(saved.cloud.pending, false);
			if (action === 'copy') {
				assert.equal(saved.cloud.file.itemId, 'copy');
				assert.deepEqual(saved.bytes, bytes('local'));
			} else {
				assert.deepEqual(adopted.bytes, bytes('remote'));
				const backup = (await library.list()).find((item) => !item.cloud);
				assert.deepEqual(backup.bytes, bytes('local'));
			}
		}
		assert.deepEqual(remote.bytes, bytes('remote'));
	});
}

test('existing-link sharing keeps grants unchanged; new access requires explicit options', async () => {
	const { cloud, client } = harness('vsdx');
	const doc = await cloud.open('link');
	let grants = 0;
	client.share = async (_file, options) => {
		grants++;
		assert.equal(options.scope, 'organization');
		return 'new-link';
	};
	assert.equal(await cloud.share(doc, { type: 'existing' }), file().webUrl);
	assert.equal(grants, 0);
	assert.equal(await cloud.share(doc, { type: 'view', scope: 'organization' }), 'new-link');
	assert.equal(grants, 1);
	doc.cloud.pending = true;
	await assert.rejects(cloud.share(doc, { type: 'edit', scope: 'anonymous' }), /resolve/);
});

test('unconfigured hosts fail honestly and identity serialization excludes capabilities', async () => {
	const cloud = new SuiteOneDrive({ store: store() });
	await assert.rejects(cloud.open('link'), /not configured/);
	assert.equal(cloud.configured(), false);
	assert.equal('webUrl' in cloudIdentity(file()), false);
	assert.equal('accessToken' in cloudIdentity({ ...file(), accessToken: 'secret' }), false);
});

test('custom providers persist their identity, isolate connections and receive session policies', async () => {
	const library = store();
	const safe = { key: 'report', revision: 1, name: 'Report.xlsx', kind: 'xlsx', access: 'write' };
	const opened = { file: { ...safe, temporaryToken: 'never persist' }, bytes: bytes('base') };
	let settings;
	const cloud = new SuiteCloudStorage({
		id: 'company-files',
		label: 'Company Files',
		store: library,
		client: { open: async () => opened, read: async () => opened },
		keyOf: (file) => file.key,
		serializeFile: ({ key, revision, name, kind, access }) => ({
			key,
			revision,
			name,
			kind,
			access,
		}),
		sessionOptions: { maxMergeAttempts: 2 },
		createSession: (_client, _opened, options) => {
			settings = options;
			return { save: async () => safe };
		},
	});
	let first = await cloud.open('company:report');
	assert.equal(first.cloud.providerId, 'company-files');
	assert.deepEqual(first.cloud.file, safe);
	assert.equal(cloudSaveLabel(first), 'Saved to Company Files');
	assert.equal((await cloud.open('company:report')).id, first.id);
	first = await library.get(first.id);
	first.cloud.pending = true;
	await cloud.save(first);
	assert.equal(settings.maxMergeAttempts, 2);
	assert.equal(cloudSaveLabel(first), 'Saved on this device; Company Files save pending');
	const other = new SuiteCloudStorage({
		id: 'different-provider',
		store: library,
		client: {},
		createSession: () => {
			throw new Error('must not connect');
		},
	});
	await assert.rejects(other.save(first), /provider that owns/);
	await assert.rejects(other.share(first, { type: 'existing' }), /provider that owns/);
	await assert.rejects(
		cloud.share({ ...first, cloud: { ...first.cloud, pending: false } }, { type: 'view' }),
		/does not support/,
	);
});

test('a cloud file is found through the bytes-free index and reuses its library document', async () => {
	const { library, cloud } = harness('docx');
	let listed = 0;
	const index = library.list;
	library.list = async () => {
		listed++;
		return index();
	};
	library.listCloud = async () =>
		(await index()).map(({ id, name, kind, revision, cloud }) => ({
			id,
			name,
			kind,
			revision,
			cloud,
		}));
	const first = await cloud.open('link');
	const second = await cloud.open('link');
	assert.equal(second.id, first.id);
	assert.equal(listed, 0, 'opening never loads every stored document');
});
