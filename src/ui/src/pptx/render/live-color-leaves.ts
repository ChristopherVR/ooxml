/**
 * The leaves a live colour drag writes, and how they are put back (see `live-color-input.ts`).
 *
 * A colour patch writes whole top-level entries (`{ shapeStyle: { ...style, fillColor } }`),
 * but the drag owns only the leaves inside them that changed. Those are found by a structural
 * diff of plain objects and arrays, each remembered with the value it had before the drag,
 * and put back onto the target as it is at that moment, so everything else in the same entry
 * keeps whatever it holds now. When a preview creates an object the target did not have (a
 * text colour on an element with no `textStyle`), the drag owns the entries it put inside, not
 * the object: putting it back removes those, and the object itself only if nothing else went
 * into it since.
 *
 * @module render/live-color-leaves
 */

export type LiveColorPath = readonly string[];

/** One leaf a drag's previews wrote, with what the target held there before the first one. */
export interface LiveColorLeaf {
	readonly path: LiveColorPath;
	readonly present: boolean;
	readonly value: unknown;
	/** An object the previews created: it goes again once putting the drag back leaves it empty. */
	readonly created?: boolean;
}

interface Diff {
	readonly changed: LiveColorPath[];
	readonly created: LiveColorPath[];
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

/** The leaves where `before` and `after` differ below `path`, and the objects `after` creates. */
function diffLeaves(
	before: unknown,
	beforeHas: boolean,
	after: unknown,
	afterHas: boolean,
	path: LiveColorPath,
	out: Diff,
): void {
	if (beforeHas === afterHas && Object.is(before, after)) {
		return;
	}
	if (before === undefined && isContainer(after)) {
		// A new object: its entries are the writes. An entry it holds as `undefined` (a cleared
		// theme ref) is one too, so putting it back removes it.
		out.created.push(path);
		for (const key of Object.keys(after)) {
			diffLeaves(undefined, false, childOf(after, key), true, [...path, key], out);
		}
		return;
	}
	if (isContainer(before) && isContainer(after) && Array.isArray(before) === Array.isArray(after)) {
		const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
		for (const key of keys) {
			diffLeaves(
				childOf(before, key),
				hasChild(before, key),
				childOf(after, key),
				hasChild(after, key),
				[...path, key],
				out,
			);
		}
		return;
	}
	out.changed.push(path);
}

const isAncestor = (ancestor: LiveColorPath, path: LiveColorPath): boolean =>
	ancestor.every((segment, index) => path[index] === segment);

/** Note every leaf `patch` writes on `value` that the drag has not written yet. */
export function recordLiveColorLeaves(
	leaves: Map<string, LiveColorLeaf>,
	value: unknown,
	patch: object,
): void {
	const diff: Diff = { changed: [], created: [] };
	for (const [key, next] of Object.entries(patch)) {
		diffLeaves(childOf(value, key), hasChild(value, key), next, true, [key], diff);
	}
	for (const path of diff.created) {
		const id = `created:${JSON.stringify(path)}`;
		if (!leaves.has(id)) {
			leaves.set(id, { path, present: false, value: undefined, created: true });
		}
	}
	for (const path of diff.changed) {
		const covered = [...leaves.values()].some(
			(leaf) => !leaf.created && isAncestor(leaf.path, path),
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

function isEmptyContainer(value: unknown): boolean {
	return isContainer(value) && Object.keys(value).length === 0;
}

/**
 * The patch that puts every leaf a drag wrote back to what it was before the drag, rebased on
 * `value` (the target now): the top-level entries it names are `value`'s own, with only those
 * leaves changed. An object the drag created goes again if that leaves it empty; a top-level
 * entry the target did not have comes back as `undefined`.
 */
export function liveColorRevertPatch<P extends object>(
	value: unknown,
	leaves: Iterable<LiveColorLeaf>,
): P {
	const patch: Record<string, unknown> = {};
	const entry = (head: string): unknown =>
		Object.hasOwn(patch, head) ? patch[head] : childOf(value, head);
	const all = [...leaves];
	for (const leaf of all.filter((candidate) => !candidate.created)) {
		const [head, ...rest] = leaf.path;
		if (head === undefined) {
			continue;
		}
		patch[head] =
			rest.length === 0
				? leaf.present
					? leaf.value
					: undefined
				: writeLeaf(entry(head), rest, leaf);
	}
	// Deepest first, so an object emptied by removing a created child goes too.
	const created = all
		.filter((candidate) => candidate.created)
		.sort((left, right) => right.path.length - left.path.length);
	for (const leaf of created) {
		const [head, ...rest] = leaf.path;
		if (head === undefined) {
			continue;
		}
		let held: unknown = entry(head);
		for (const segment of rest) {
			held = childOf(held, segment);
		}
		if (isEmptyContainer(held)) {
			patch[head] = rest.length === 0 ? undefined : writeLeaf(entry(head), rest, leaf);
		}
	}
	return patch as P;
}
