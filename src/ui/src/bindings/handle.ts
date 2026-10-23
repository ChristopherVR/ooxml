// The imperative handle every framework exposes (React ref, Vue template ref, Angular component
// instance, Solid `editorRef`, Svelte exports), built once from a spec of its members.

/**
 * How one handle member behaves while nothing is mounted:
 * - `get`: a property; reading it throws (or gives `undefined` with `lenientGetters`).
 * - `call`: a method; calling it throws.
 * - `async`: a method returning a promise; calling it rejects.
 * - `optional-call`: a method that does nothing (`markClean` before mount is harmless).
 * - `{ fallback }`: a property that reads `fallback` (`dirty` is `false`).
 */
export type HandleMember =
	| 'get'
	| 'call'
	| 'async'
	| 'optional-call'
	| { readonly fallback: unknown };

/** One entry per member of the handle `H`; a missing member is a compile error. */
export type HandleSpec<H> = { readonly [K in keyof H]-?: HandleMember };

export interface DeferredHandleOptions {
	/** The error thrown before mount and after unmount. Default `Editor is not mounted`. */
	readonly message?: string;
	/** `get` members read `undefined` instead of throwing (Vue templates read refs before mount). */
	readonly lenientGetters?: boolean;
}

/** The member names of a handle spec, in declaration order. */
export function handleKeys<H>(spec: HandleSpec<H>): (keyof H & string)[] {
	return Object.keys(spec) as (keyof H & string)[];
}

type Target = Record<string, unknown> | null | undefined;

function descriptor(
	key: string,
	member: HandleMember,
	target: (self: unknown) => Target,
	options: DeferredHandleOptions,
): PropertyDescriptor {
	const message = options.message ?? 'Editor is not mounted';
	const required = (self: unknown): Record<string, unknown> => {
		const current = target(self);
		if (!current) throw new Error(message);
		return current;
	};
	const invoke = (current: Record<string, unknown>, args: unknown[]): unknown =>
		(current[key] as (...values: unknown[]) => unknown)(...args);
	if (typeof member === 'object')
		return {
			get(this: unknown) {
				const current = target(this);
				return current ? current[key] : member.fallback;
			},
			enumerable: true,
			configurable: true,
		};
	if (member === 'get')
		return {
			get(this: unknown) {
				return options.lenientGetters ? target(this)?.[key] : required(this)[key];
			},
			enumerable: true,
			configurable: true,
		};
	const value =
		member === 'async'
			? async function (this: unknown, ...args: unknown[]) {
					return invoke(required(this), args);
				}
			: member === 'optional-call'
				? function (this: unknown, ...args: unknown[]) {
						const current = target(this);
						return current ? invoke(current, args) : undefined;
					}
				: function (this: unknown, ...args: unknown[]) {
						return invoke(required(this), args);
					};
	return { value, enumerable: true, configurable: true, writable: true };
}

/**
 * A handle for adapters whose binding may not be mounted yet: every member reads the current
 * binding, so one handle object serves the whole lifetime of the component.
 */
export function deferredHandle<H>(
	spec: HandleSpec<H>,
	current: () => H | null | undefined,
	options: DeferredHandleOptions = {},
): H {
	const handle = {};
	const target = (): Target => current() as Target;
	for (const key of handleKeys(spec))
		Object.defineProperty(handle, key, descriptor(key, spec[key], target, options));
	return handle as H;
}

/**
 * Defines the handle members on a class prototype (an Angular component), reading the binding of
 * each instance through `current`. Declare the members on the class with an interface merge.
 */
export function installHandle<H, T>(
	prototype: T,
	spec: HandleSpec<H>,
	current: (self: T) => H | null | undefined,
	options: DeferredHandleOptions = {},
): void {
	const target = (self: unknown): Target => current(self as T) as Target;
	for (const key of handleKeys(spec))
		Object.defineProperty(prototype, key, descriptor(key, spec[key], target, options));
}
