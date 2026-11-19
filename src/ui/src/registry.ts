/** Internal UI ABI, independent of package versions. Bump for incompatible contracts. */
export const CONTRACT_REVISION = 1;
const CONTRACT_KEY = Symbol.for('office-ui.web-control-contract');

type Stamped = CustomElementConstructor & { [CONTRACT_KEY]?: number };

/** Throws when a tag is already defined by an incompatible ooxml-ui build. */
export function assertContract(registry: Pick<CustomElementRegistry, 'get'>, tag: string): void {
	const revision = (registry.get(tag) as Stamped | undefined)?.[CONTRACT_KEY];
	if (revision !== undefined && revision !== CONTRACT_REVISION) {
		throw new Error(
			`Incompatible ${tag} contract (${revision}; expected ${CONTRACT_REVISION}). Load one version of ooxml-ui per window.`,
		);
	}
}

/**
 * Define `tag` once per registry. Classes are built lazily inside `make`, so importing this
 * package never touches `HTMLElement` (SSR-safe); only a define call needs a DOM.
 */
export function defineOnce(
	registry: CustomElementRegistry,
	tag: string,
	make: () => CustomElementConstructor,
): void {
	assertContract(registry, tag);
	if (registry.get(tag)) return;
	defineElement(registry, tag, make());
}

/**
 * The one way an ooxml-ui class enters a registry: it carries the contract stamp and honours
 * properties a host set before the definition loaded. The caller has checked the tag is free.
 */
export function defineElement(
	registry: CustomElementRegistry,
	tag: string,
	ctor: CustomElementConstructor,
): void {
	markContract(ctor);
	adoptEarlyProperties(ctor);
	registry.define(tag, ctor);
}

/** Stamp a constructor with this build's contract revision. */
export function markContract(ctor: CustomElementConstructor): void {
	Object.defineProperty(ctor, CONTRACT_KEY, { value: CONTRACT_REVISION });
}

const adopting = new WeakSet<object>();
const settersOf = new WeakMap<object, ReadonlySet<string>>();

/** Every property the element's own classes implement with a setter, above the platform's. */
function classSetters(element: HTMLElement): ReadonlySet<string> {
	const start = Object.getPrototypeOf(element) as object;
	let names = settersOf.get(start);
	if (names) return names;
	const platform = element.ownerDocument.defaultView?.HTMLElement.prototype;
	const found = new Set<string>();
	for (
		let proto: object | null = start;
		proto && proto !== platform;
		proto = Object.getPrototypeOf(proto) as object | null
	)
		for (const name of Object.getOwnPropertyNames(proto))
			if (Object.getOwnPropertyDescriptor(proto, name)?.set) found.add(name);
	names = found;
	settersOf.set(start, names);
	return names;
}

const heldEarly = new WeakMap<HTMLElement, [string, unknown][]>();

/**
 * A value assigned before the definition loaded is an own property hiding the class's accessor.
 * Lifts those values off the element, in assignment order, so that its constructor, first render
 * and connection all see the class's defaults. Safe to call more than once.
 */
export function holdEarlyProperties(element: HTMLElement): void {
	const setters = classSetters(element);
	if (!setters.size) return;
	const target = element as unknown as Record<string, unknown>;
	for (const name of Object.keys(element)) {
		if (!setters.has(name)) continue;
		const held = heldEarly.get(element) ?? [];
		held.push([name, target[name]]);
		heldEarly.set(element, held);
		delete target[name];
	}
}

/**
 * Gives held values to the accessors once the element is connected and its shell exists, the path
 * of any later write. A setter that refuses its value is reported and stops nothing else.
 */
function releaseEarlyProperties(element: HTMLElement): void {
	const held = heldEarly.get(element);
	if (!held) return;
	heldEarly.delete(element);
	const target = element as unknown as Record<string, unknown>;
	for (const [name, value] of held) {
		try {
			target[name] = value;
		} catch (error) {
			queueMicrotask(() => {
				throw error;
			});
		}
	}
}

/** Wraps each connection of `ctor`'s elements so early properties take effect. */
function adoptEarlyProperties(ctor: CustomElementConstructor): void {
	const proto = ctor.prototype as { connectedCallback?: () => void };
	if (adopting.has(proto)) return;
	adopting.add(proto);
	const connected = proto.connectedCallback;
	Object.defineProperty(proto, 'connectedCallback', {
		configurable: true,
		writable: true,
		value(this: HTMLElement): void {
			holdEarlyProperties(this);
			try {
				connected?.call(this);
			} finally {
				releaseEarlyProperties(this);
			}
		},
	});
}

export type Definer = ((registry?: CustomElementRegistry) => void) & { readonly tag: string };

/** The registry of the current window, or undefined where there is no DOM (SSR). */
export function browserRegistry(): CustomElementRegistry | undefined {
	return typeof window === 'undefined' ? undefined : window.customElements;
}

/**
 * Wrap a per-tag `make` factory into a public, SSR-safe, idempotent definer. `requires` are the
 * elements this one renders inside its own template; defining the composite defines them first, so
 * a product that registers only the element it uses never gets a silently inert inner control.
 */
export function definer(
	tag: string,
	make: () => CustomElementConstructor,
	requires: readonly Definer[] = [],
): Definer {
	const define = (registry: CustomElementRegistry | undefined = browserRegistry()): void => {
		if (!registry) return;
		for (const dependency of requires) dependency(registry);
		defineOnce(registry, tag, make);
	};
	return Object.assign(define, { tag });
}

/** Bubbling, composed event: crosses shadow boundaries like the native ones do. */
export function emit<T>(host: HTMLElement, type: string, detail: T, cancelable = false): boolean {
	return host.dispatchEvent(
		new CustomEvent<T>(type, { detail, bubbles: true, composed: true, cancelable }),
	);
}

/**
 * A boolean property value as its attribute would read it: `''` means present, as in
 * `checked=""`. Frameworks that pass attribute idioms through properties (Svelte's
 * `checked={on ? '' : undefined}`) therefore keep working.
 */
export function present(value: unknown): boolean {
	return value === '' || Boolean(value);
}
