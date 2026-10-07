import { expect, it } from 'vitest';
import { createAssetSync } from './assets.js';

it('routes raw binary payloads without aliasing callers or incrementing equal swaps', () => {
	const sync = createAssetSync({ mapName: 'assets', fields: { bytes: '_bytes' } });
	const owner = new Map<string, unknown>();
	const assets = new Map<string, unknown>();
	const bytes = new Uint8Array([1, 2, 3]);
	sync.write('image', { bytes }, owner, assets);
	bytes[0] = 9;
	const target: Record<string, unknown> = {};
	sync.read(owner, assets, target);
	expect(target.bytes).toEqual(new Uint8Array([1, 2, 3]));
	(target.bytes as Uint8Array)[0] = 9;
	sync.reconcile('image', { bytes: new Uint8Array([1, 2, 3]) }, owner, assets);
	expect(owner.has('bytes__v')).toBe(false);
	sync.reconcile('image', { bytes: new Uint8Array([4, 5, 6]) }, owner, assets);
	expect(owner.get('bytes__v')).toBe(1);
	sync.read(owner, assets, target);
	expect(target.bytes).toEqual(new Uint8Array([4, 5, 6]));
	sync.reconcile('image', {}, owner, assets);
	expect(owner.has('_bytes')).toBe(false);
	expect(assets.size).toBe(0);
});

it('retains the existing string payload and version-counter contract', () => {
	const sync = createAssetSync({ mapName: 'assets', fields: { mediaData: '_mdRef' } });
	const owner = new Map<string, unknown>();
	const assets = new Map<string, unknown>();
	sync.write('shape', { mediaData: 'data:image/png;base64,AAA' }, owner, assets);
	sync.reconcile('shape', { mediaData: 'data:image/png;base64,AAA' }, owner, assets);
	expect(owner.has('mediaData__v')).toBe(false);
	const target: Record<string, unknown> = {};
	sync.read(owner, assets, target);
	expect(target.mediaData).toBe('data:image/png;base64,AAA');
	sync.reconcile('shape', { mediaData: 'data:image/png;base64,BBB' }, owner, assets);
	expect(owner.get('mediaData__v')).toBe(1);
});
