// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Style inheritance (`w:basedOn`). A chain can be thousands of styles long and can loop, so each
// catalog's chains are resolved once, in time linear in the catalog, never walked again per
// paragraph or run. Catalogs are immutable values (an edit builds a new `styles` record), which is
// what makes keeping the results by catalog safe.

/** A style that may be based on another. */
export interface Inheriting {
	readonly basedOn?: string | undefined;
}

export type StyleRecord<T extends Inheriting> = Readonly<Record<string, T>>;

const parentOf = <T extends Inheriting>(styles: StyleRecord<T>, id: string): string | undefined => {
	const parent = styles[id]?.basedOn;
	return parent !== undefined && Object.hasOwn(styles, parent) ? parent : undefined;
};

/**
 * Where each style's chain first reaches a style it already passed, in catalog order: the style
 * itself when it lies on a cycle, else the cycle member its chain enters. Styles whose chain ends
 * are left out, and each entry is listed once.
 */
export function inheritanceCycleEntries<T extends Inheriting>(styles: StyleRecord<T>): string[] {
	const entry = new Map<string, string | undefined>();
	for (const id of Object.keys(styles)) {
		const path: string[] = [];
		const onPath = new Map<string, number>();
		let current: string | undefined = id;
		while (current !== undefined && !entry.has(current) && !onPath.has(current)) {
			onPath.set(current, path.length);
			path.push(current);
			current = parentOf(styles, current);
		}
		const start = current === undefined ? undefined : onPath.get(current);
		if (start === undefined) {
			const reached = current === undefined ? undefined : entry.get(current);
			for (const step of path) entry.set(step, reached);
		} else {
			// `path[start..]` is a cycle; the styles before it lead into it at `current`.
			path.forEach((step, index) => entry.set(step, index >= start ? step : current));
		}
	}
	const entries = new Set<string>();
	for (const id of Object.keys(styles)) {
		const reached = entry.get(id);
		if (reached !== undefined) entries.add(reached);
	}
	return [...entries];
}

/**
 * Resolves a property set along each style's chain, nearest style last applied (so it wins), as
 * `merge(merge(merge(initial, root), ...), style)`. `own` is a style's contribution, `merge` adds a
 * contribution to a copy of the result so far. A style on a cycle takes every style of its cycle,
 * from the one after it round to itself, as a walk that stops at the first repeat does.
 */
export function chainResolver<T extends Inheriting, R extends object>(
	styles: StyleRecord<T>,
	own: (style: T, id: string) => R | undefined,
	merge: (resolved: R, contribution: R) => R,
	initial: R,
): (id: string | undefined) => R | undefined {
	const resolved = new Map<string, R>();
	const add = (base: R, id: string): R => {
		const contribution = own(styles[id]!, id);
		return contribution ? merge(base, contribution) : base;
	};
	const resolveCycle = (members: string[]): void => {
		// Two laps backwards: member i ends with every member from i + c - 1 down to i applied.
		let acc = initial;
		for (let lap = 0; lap < 2; lap++)
			for (let index = members.length - 1; index >= 0; index--) {
				const id = members[index]!;
				acc = add(acc, id);
				if (lap === 1) resolved.set(id, acc);
			}
	};
	const resolve = (id: string): R => {
		const path: string[] = [];
		const onPath = new Map<string, number>();
		let current: string | undefined = id;
		while (current !== undefined && !resolved.has(current) && !onPath.has(current)) {
			onPath.set(current, path.length);
			path.push(current);
			current = parentOf(styles, current);
		}
		let end = path.length;
		const cycleStart = current === undefined ? undefined : onPath.get(current);
		if (cycleStart !== undefined) {
			resolveCycle(path.slice(cycleStart));
			end = cycleStart;
		}
		let acc = current === undefined ? initial : resolved.get(current)!;
		for (let index = end - 1; index >= 0; index--) {
			const step = path[index]!;
			acc = add(acc, step);
			resolved.set(step, acc);
		}
		return resolved.get(id)!;
	};
	return (id) => (id !== undefined && Object.hasOwn(styles, id) ? resolve(id) : undefined);
}

/** `chainResolver`s kept per catalog `styles` record and per `kind` of resolution. */
export function cachedResolver<T extends Inheriting, R extends object>(
	cache: WeakMap<object, (id: string | undefined) => R | undefined>,
	styles: StyleRecord<T>,
	create: () => (id: string | undefined) => R | undefined,
): (id: string | undefined) => R | undefined {
	let resolver = cache.get(styles);
	if (!resolver) {
		resolver = create();
		cache.set(styles, resolver);
	}
	return resolver;
}

const defaults = new WeakMap<object, string | undefined>();

/** The id of the catalog's default style (`w:default="1"`), found once per catalog. */
export function defaultStyleId<T extends { readonly id: string; readonly isDefault?: boolean }>(
	styles: Readonly<Record<string, T>>,
): string | undefined {
	if (!defaults.has(styles))
		defaults.set(styles, Object.values(styles).find((style) => style.isDefault)?.id);
	return defaults.get(styles);
}

/** Copies the properties `contribution` defines over `resolved` (an explicit `undefined` is skipped). */
export function mergeDefined<R extends object>(resolved: R, contribution: R): R {
	const next = { ...resolved } as Record<string, unknown>;
	for (const [key, value] of Object.entries(contribution))
		if (value !== undefined) next[key] = value;
	return next as R;
}
