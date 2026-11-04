import * as Y from 'yjs';

const record = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' &&
	value !== null &&
	(Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);

/** Snapshot equality, independent of plain-object key insertion order. Arrays remain ordered. */
export function snapshotEqual(a: unknown, b: unknown): boolean {
	if (Object.is(a, b)) return true;
	if (a instanceof Uint8Array && b instanceof Uint8Array)
		return a.length === b.length && a.every((value, index) => value === b[index]);
	if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
	if (a instanceof Map && b instanceof Map)
		return (
			a.size === b.size &&
			[...a].every(([key, value]) => b.has(key) && snapshotEqual(value, b.get(key)))
		);
	if (Array.isArray(a) && Array.isArray(b))
		return a.length === b.length && a.every((value, index) => snapshotEqual(value, b[index]));
	if (!record(a) || !record(b)) return false;
	const ak = Object.keys(a).filter((key) => a[key] !== undefined);
	const bk = Object.keys(b).filter((key) => b[key] !== undefined);
	return ak.length === bk.length && ak.every((key) => snapshotEqual(a[key], b[key]));
}

/**
 * Reconciles saved snapshots, not invented CRDT history. Independent object fields combine;
 * conflicting values, ordered arrays and delete-versus-edit remain explicitly unresolved.
 */
export function mergeSnapshotValue(
	base: unknown,
	local: unknown,
	remote: unknown,
	path: string,
	issues: string[],
): unknown {
	if (snapshotEqual(local, remote)) return structuredClone(local);
	if (snapshotEqual(local, base)) return structuredClone(remote);
	if (snapshotEqual(remote, base)) return structuredClone(local);
	if (base instanceof Map && local instanceof Map && remote instanceof Map) {
		const out = new Map<unknown, unknown>();
		for (const key of new Set([...base.keys(), ...local.keys(), ...remote.keys()])) {
			if (typeof key !== 'string' && typeof key !== 'number') {
				issues.push(`Unsupported map identity at ${path}.`);
				continue;
			}
			const value = mergeSnapshotValue(
				base.get(key),
				local.get(key),
				remote.get(key),
				`${path}/${key}`,
				issues,
			);
			if (value !== undefined) out.set(key, value);
		}
		return out;
	}
	if ((base === undefined || record(base)) && record(local) && record(remote)) {
		const out: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
		for (const key of new Set([
			...Object.keys(base ?? {}),
			...Object.keys(local),
			...Object.keys(remote),
		])) {
			const value = mergeSnapshotValue(
				base?.[key],
				local[key],
				remote[key],
				`${path}/${key}`,
				issues,
			);
			if (value !== undefined) out[key] = value;
		}
		return out;
	}
	issues.push(`Concurrent changes overlap at ${path}.`);
	return undefined;
}

function assignMap(
	target: Y.Map<unknown>,
	values: Record<string, unknown>,
	templates: Y.Map<unknown>[],
): void {
	for (const key of [...target.keys()]) if (!(key in values)) target.delete(key);
	for (const [key, value] of Object.entries(values)) {
		const children = templates
			.map((map) => map.get(key))
			.filter((item): item is Y.Map<unknown> => item instanceof Y.Map);
		if (children.length && record(value)) {
			let child = target.get(key);
			if (!(child instanceof Y.Map)) {
				child = new Y.Map<unknown>();
				target.set(key, child);
			}
			assignMap(child as Y.Map<unknown>, value, children);
		} else target.set(key, structuredClone(value));
	}
}

/** Uses real common-base Yjs updates, then reconciles overlapping atomic JSON values explicitly. */
export function mergeSnapshotMaps(
	base: Y.Doc,
	local: Y.Doc,
	remote: Y.Doc,
	names: readonly string[],
): { doc: Y.Doc; issues: string[] } {
	const issues: string[] = [];
	const values = names.map((name) =>
		mergeSnapshotValue(
			base.getMap(name).toJSON(),
			local.getMap(name).toJSON(),
			remote.getMap(name).toJSON(),
			name,
			issues,
		),
	);
	const doc = new Y.Doc();
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(base));
	if (!issues.length) {
		const vector = Y.encodeStateVector(base);
		Y.applyUpdate(doc, Y.encodeStateAsUpdate(local, vector));
		Y.applyUpdate(doc, Y.encodeStateAsUpdate(remote, vector));
		doc.transact(() =>
			names.forEach((name, index) =>
				assignMap(doc.getMap(name), values[index] as Record<string, unknown>, [
					base.getMap(name),
					local.getMap(name),
					remote.getMap(name),
				]),
			),
		);
	}
	return { doc, issues };
}
