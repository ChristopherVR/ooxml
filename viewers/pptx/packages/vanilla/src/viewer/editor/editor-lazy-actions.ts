/**
 * Build a lazily-resolved proxy over an action bag whose real instance isn't
 * constructed yet at call time (`EditActions` is built by `createEditActions`
 * *after* the chrome/ribbon that needs to reference its methods, see
 * `PptxViewer.ts`'s construction order). Every method call forwards to
 * `getActions()` at the moment it's invoked, so the ribbon can be wired to a
 * stable object at construction time while the underlying `EditActions`
 * instance is swapped in once the editor controller exists (and again on
 * `setLocale`'s chrome remount).
 *
 * Nested action groups (`EditActions.sections`, `.comments`, `.masterView`)
 * resolve the same way: `lazy.sections.addSection(...)` forwards to
 * `getActions().sections.addSection(...)` at call time.
 */
export function createLazyActions<T extends object>(getActions: () => T): T {
	return new Proxy<T>({} as T, { get: lazyMembers(getActions) });
}

type Callable = (...args: unknown[]) => unknown;
type Bag = Record<string, unknown>;

/** A `get` trap that hands out one cached {@link lazyMember} per property name. */
function lazyMembers(resolve: () => object): (target: object, prop: string | symbol) => unknown {
	const cache = new Map<string, Callable>();
	return (_target, prop) => {
		if (typeof prop !== 'string') {
			return undefined;
		}
		let member = cache.get(prop);
		if (!member) {
			member = lazyMember(resolve, prop);
			cache.set(prop, member);
		}
		return member;
	};
}

/**
 * The forwarder for `resolve()[prop]`: calling it calls that method on the
 * current owner (so `this` stays the owner), and reading a property from it
 * descends into that member as a nested action group.
 */
function lazyMember(resolve: () => object, prop: string): Callable {
	const owner = (): Bag => resolve() as Bag;
	return new Proxy<Callable>((() => undefined) as Callable, {
		apply: (_target, _this, args: unknown[]) => (owner()[prop] as Callable).apply(owner(), args),
		get: lazyMembers(() => owner()[prop] as object),
	});
}
