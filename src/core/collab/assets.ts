// Binary payload sync. Large binary fields (media, OLE, embedded images, 3D models) are synced once
// through a separate root Y.Map keyed by `${ownerId}:${field}`, instead of being re-embedded as base64
// in every write of the owning object. The owner's own map keeps only a small ref pointer, plus a
// version counter that is bumped on an in-place swap so deep observers fire.
//
// Generalised from pptx-viewer `collaboration-assets.ts`: the field-to-ref-key table and the assets map
// name are now a spec supplied by the product (pptx passes its `pptx:assets` / `_mdRef` ... table, so
// its documents stay wire-compatible). Known, accepted cost carried over: entries are not
// garbage-collected when the owner is deleted outright; a long-lived room's asset map only grows.
// See PROVENANCE.md.

/** Structural view of a `Y.Map` (a real `Y.Map` satisfies it). */
export interface YMapLike {
	get: (key: string) => unknown;
	set: (key: string, value: unknown) => void;
	delete: (key: string) => void;
	forEach: (cb: (value: unknown, key: string) => void) => void;
}

export interface AssetSpec {
	/** Name of the root Y.Map holding the payloads (for example `pptx:assets`). */
	mapName: string;
	/** Field name to the ref-pointer key stored on the owner's map (`mediaData` to `_mdRef`). */
	fields: Readonly<Record<string, string>>;
	/** Suffix of the per-field version counter; defaults to `__v`. */
	versionSuffix?: string;
}

export interface AssetSync {
	readonly spec: AssetSpec;
	/** Fields routed through the asset map. */
	isAssetField: (field: string) => boolean;
	/** Whether `key` is one of the ref-pointer keys this spec owns. */
	isRefKey: (key: string) => boolean;
	/** Whether `key` is an internal version counter (never a model field). */
	isVersionKey: (key: string) => boolean;
	assetKey: (ownerId: string, field: string) => string;
	getMap: (doc: { getMap: (name: string) => YMapLike }) => YMapLike;
	/**
	 * Write present asset fields of `record` into `assets`, storing only a ref on `owner`. Never reads
	 * `owner`, so it is safe on a brand-new, not yet integrated map (Yjs throws on reading one).
	 */
	write: (
		ownerId: string,
		record: Record<string, unknown>,
		owner: YMapLike,
		assets: YMapLike,
	) => void;
	/**
	 * Variant for an already integrated owner: skips redundant sets, bumps the version counter on an
	 * in-place payload swap, and clears the ref and the payload when the field was removed.
	 */
	reconcile: (
		ownerId: string,
		record: Record<string, unknown>,
		owner: YMapLike,
		assets: YMapLike,
	) => void;
	/** Rehydrate the asset fields referenced by `owner` onto `target`. */
	read: (owner: YMapLike, assets: YMapLike, target: Record<string, unknown>) => void;
}

const isPayload = (value: unknown): value is string =>
	typeof value === 'string' && value.length > 0;

export function createAssetSync(spec: AssetSpec): AssetSync {
	const suffix = spec.versionSuffix ?? '__v';
	const fieldNames = Object.keys(spec.fields);
	const refToField = new Map(Object.entries(spec.fields).map(([field, ref]) => [ref, field]));
	const assetKey = (ownerId: string, field: string): string => `${ownerId}:${field}`;
	const versionKey = (field: string): string => `${field}${suffix}`;
	const refOf = (field: string): string => spec.fields[field]!;

	return {
		spec,
		isAssetField: (field) => field in spec.fields,
		isRefKey: (key) => refToField.has(key),
		isVersionKey: (key) => key.endsWith(suffix) && key.slice(0, -suffix.length) in spec.fields,
		assetKey,
		getMap: (doc) => doc.getMap(spec.mapName),
		write(ownerId, record, owner, assets) {
			for (const field of fieldNames) {
				const value = record[field];
				if (!isPayload(value)) continue;
				const key = assetKey(ownerId, field);
				if (assets.get(key) !== value) assets.set(key, value);
				owner.set(refOf(field), key);
			}
		},
		reconcile(ownerId, record, owner, assets) {
			for (const field of fieldNames) {
				const refKey = refOf(field);
				const value = record[field];
				if (isPayload(value)) {
					const key = assetKey(ownerId, field);
					if (assets.get(key) !== value) {
						assets.set(key, value);
						// An in-place swap leaves the ref unchanged, so nothing on the owner would move
						// and no peer would re-read the payload: bump a counter to force a transaction.
						const current = owner.get(versionKey(field));
						owner.set(versionKey(field), (typeof current === 'number' ? current : 0) + 1);
					}
					if (owner.get(refKey) !== key) owner.set(refKey, key);
				} else {
					const existing = owner.get(refKey);
					if (typeof existing === 'string') {
						owner.delete(refKey);
						assets.delete(existing);
					}
				}
			}
		},
		read(owner, assets, target) {
			for (const [refKey, field] of refToField) {
				const ref = owner.get(refKey);
				if (typeof ref !== 'string') continue;
				const value = assets.get(ref);
				if (typeof value === 'string') target[field] = value;
			}
		},
	};
}
