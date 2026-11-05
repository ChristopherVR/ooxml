import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveCloudOptions } from './cloud-options.js';

const input = { kind: 'xlsx', base: {}, local: new Uint8Array(), remote: {} };
test('OneDrive hosts can replace the client, disable merging and set rebase limits', () => {
	const client = {};
	const options = resolveCloudOptions({
		ooxmlOneDriveOptions: { client, merge: false, maxMergeAttempts: 2 },
	});
	assert.equal(options.client, client);
	assert.equal(options.id, 'onedrive');
	assert.equal(options.merge, undefined);
	assert.equal(options.sessionOptions.maxMergeAttempts, 2);
});
test('per-format handlers can be disabled or replaced without loading an Office engine', async () => {
	const options = resolveCloudOptions({
		ooxmlOneDriveOptions: { client: {}, mergers: { xlsx: undefined } },
	});
	assert.deepEqual(await options.merge(input), {
		status: 'unresolved',
		issues: ['unsupported-merge-kind:xlsx'],
	});
	const override = async () => ({ status: 'merged', bytes: new Uint8Array([1]) });
	const custom = resolveCloudOptions({
		ooxmlOneDriveOptions: { client: {}, mergers: { xlsx: override } },
	});
	assert.deepEqual(await custom.merge(input), { status: 'merged', bytes: new Uint8Array([1]) });
});
test('custom providers take precedence and preserve host policy callbacks', () => {
	const custom = {
		id: 'custom',
		label: 'Custom',
		client: {},
		keyOf: (file) => file.key,
		serializeFile: (file) => file,
		merge: () => ({ status: 'unresolved', issues: [] }),
		resolveConflict: () => 'cancel',
		shareChoices: [],
		createSession: () => {},
	};
	const options = resolveCloudOptions({
		ooxmlCloudOptions: custom,
		ooxmlOneDriveOptions: { client: {} },
	});
	assert.equal(options.client, custom.client);
	assert.equal(options.merge, custom.merge);
	assert.equal(options.resolveConflict, custom.resolveConflict);
	assert.equal(options.createSession, custom.createSession);
	assert.equal(options.shareChoices, custom.shareChoices);
	assert.equal(options.id, 'custom');
	assert.throws(() => resolveCloudOptions({ ooxmlCloudOptions: { id: 'custom' } }), TypeError);
});

test('a OneDrive host gets a browsable provider; without a sign-in hook it is already signed in', async () => {
	const none = resolveCloudOptions({});
	assert.equal(none.fileProvider, undefined);
	const options = resolveCloudOptions({
		ooxmlOneDriveOptions: { getAccessToken: async () => 'token' },
	});
	assert.equal(options.fileProvider.account()?.name, 'Microsoft account');
	let signIns = 0;
	const hosted = resolveCloudOptions({
		ooxmlOneDriveOptions: {
			getAccessToken: async () => 'token',
			signIn: async () => {
				signIns++;
				return { name: 'Ada' };
			},
		},
	});
	assert.equal(hosted.fileProvider.account(), null);
	assert.equal((await hosted.fileProvider.signIn()).name, 'Ada');
	assert.equal(hosted.fileProvider.account()?.name, 'Ada');
	assert.equal(signIns, 1);
	const supplied = { id: 'x' };
	const custom = resolveCloudOptions({
		ooxmlCloudOptions: {
			id: 'company',
			label: 'Company',
			client: {},
			keyOf: (file) => file.key,
			serializeFile: (file) => file,
			fileProvider: supplied,
		},
	});
	assert.equal(custom.fileProvider, supplied);
});
