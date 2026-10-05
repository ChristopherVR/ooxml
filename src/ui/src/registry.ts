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
	const ctor = make();
	Object.defineProperty(ctor, CONTRACT_KEY, { value: CONTRACT_REVISION });
	registry.define(tag, ctor);
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
