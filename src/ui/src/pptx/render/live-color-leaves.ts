/**
 * The leaves a live colour drag writes, and how they are put back (see `live-color-input.ts`).
 *
 * A colour patch writes whole top-level entries (`{ shapeStyle: { ...style, fillColor } }`),
 * but the drag owns only the leaves inside them that changed. Those are found by a structural
 * diff of plain objects and arrays, each remembered with the value it had before the drag,
 * and put back onto the target as it is at that moment, so everything else in the same entry
 * keeps whatever it holds now.
 *
 * @module render/live-color-leaves
 */

export type LiveColorPath = readonly string[];

/** One leaf a drag's previews wrote, with what the target held there before the first one. */
export interface LiveColorLeaf {
	readonly path: LiveColorPath;
	readonly present: boolean;
	readonly value: unknown;
}

type Container = Record<string, unknown> | unknown[];

function isContainer(value: unknown): value is Container {
	if (Array.isArray(value)) {
		return true;
	}
	if (typeof value !== 'object' || value === null) {
		return false;
	}
	const proto = Object.getPrototypeOf(value) as unknown;
	return proto === Object.prototype || proto === null;
}

function childOf(value: unknown, key: string): unknown {
	return isContainer(value) ? (value as Record<string, unknown>)[key] : undefined;
}

function hasChild(value: unknown, key: string): boolean {
	return isContainer(value) && Object.hasOwn(value, key);
}

/** The leaves where `before` and `after` differ, below `path`. */
function diffLeaves(
	before: unknown,
	after: unknown,
	path: LiveColorPath,
	out: LiveColorPath[],
): void {
	if (Object.is(before, after)) {
		return;
	}
	if (isContainer(before) && isContainer(after) && Array.isArray(before) === Array.isArray(after)) {
		const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
		for (const key of keys) {
			if (hasChild(before, key) !== hasChild(after, key)) {
				// An entry the patch adds as `undefined` (a cleared theme ref) is a write too: putting
				// it back removes it.
				out.push([...path, key]);
			} else {
				diffLeaves(childOf(before, key), childOf(after, key), [...path, key], out);
			}
		}
		return;
	}
	out.push(path);
}

/** Note every leaf `patch` writes on `value` that the drag has not written yet. */
export function recordLiveColorLeaves(
	leaves: Map<string, LiveColorLeaf>,
	value: unknown,
	patch: object,
): void {
	const changed: LiveColorPath[] = [];
	for (const [key, next] of Object.entries(patch)) {
		diffLeaves(childOf(value, key), next, [key], changed);
	}
	for (const path of changed) {
		const covered = [...leaves.values()].some((leaf) =>
			leaf.path.every((segment, index) => path[index] === segment),
		);
		if (covered) {
			continue;
		}
		let holder: unknown = value;
		for (const segment of path.slice(0, -1)) {
			holder = childOf(holder, segment);
		}
		const last = path[path.length - 1]!;
		leaves.set(JSON.stringify(path), {
			path,
			present: hasChild(holder, last),
			value: childOf(holder, last),
		});
	}
}

/** `container` with `leaf` written at `path` (or removed, when the target did not have it). */
function writeLeaf(container: unknown, path: LiveColorPath, leaf: LiveColorLeaf): unknown {
	const [head, ...rest] = path;
	const copy: Container = Array.isArray(container)
		? [...container]
		: { ...(isContainer(container) ? container : {}) };
	const record = copy as Record<string, unknown>;
	if (rest.length > 0) {
		record[head!] = writeLeaf(record[head!], rest, leaf);
	} else if (leaf.present) {
		record[head!] = leaf.value;
	} else if (Array.isArray(copy)) {
		record[head!] = undefined;
	} else {
		delete record[head!];
	}
	return copy;
}

/**
 * The patch that puts every leaf a drag wrote back to what it was before the drag, rebased on
 * `value` (the target now): the top-level entries it names are `value`'s own, with only those
 * leaves changed. A top-level entry the target did not have comes back as `undefined`.
 */
export function liveColorRevertPatch<P extends object>(
	value: unknown,
	leaves: Iterable<LiveColorLeaf>,
): P {
	const patch: Record<string, unknown> = {};
	for (const leaf of leaves) {
		const [head, ...rest] = leaf.path;
		if (head === undefined) {
			continue;
		}
		if (rest.length === 0) {
			patch[head] = leaf.present ? leaf.value : undefined;
		} else {
			const base = Object.hasOwn(patch, head) ? patch[head] : childOf(value, head);
			patch[head] = writeLeaf(base, rest, leaf);
		}
	}
	return patch as P;
}
