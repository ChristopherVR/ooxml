/**
 * Properties a host assigned to an element before its definition loaded. Until the upgrade they
 * are plain own properties hiding the class's accessors. The rule, the one the standard custom
 * element pattern and Lit follow: such a value counts as written just after the upgrade, so it
 * wins over the attributes the element had then, and any later write, property or attribute, wins
 * over it.
 *
 * Held values are replayed through the class's setters when the element connects (before its own
 * `connectedCallback` runs, so mount-time options apply to the first build), or after the current
 * script when it was upgraded while detached. Until then each name has an own accessor that reads
 * the class's value and passes writes through; a write that is not part of the upgrade drops the
 * held value for that name.
 */

interface Held {
	/** Name to early value, in assignment order. */
	values: Map<string, unknown>;
	/** The attributes at the upgrade whose `attributeChangedCallback` has not run yet. */
	attributes: Map<string, string>;
}

const held = new WeakMap<HTMLElement, Held>();
/** Elements inside the `attributeChangedCallback` the upgrade makes for one of their attributes. */
const upgrading = new WeakSet<HTMLElement>();
const accessors = new WeakSet<object>();
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

/** Removes the stand-in accessor of `name`, leaving anything else (a class field) alone. */
function removeAccessor(element: HTMLElement, name: string): void {
	const own = Object.getOwnPropertyDescriptor(element, name);
	if (own?.get && accessors.has(own.get))
		delete (element as unknown as Record<string, unknown>)[name];
}

/**
 * Lifts the early values off `element` so its constructor and attribute callbacks run against the
 * class's defaults, and schedules their replay. Safe to call more than once.
 */
export function holdEarlyProperties(element: HTMLElement): void {
	if (held.has(element)) return;
	const setters = classSetters(element);
	const names = Object.keys(element).filter((name) => setters.has(name));
	if (!names.length) return;
	const target = element as unknown as Record<string, unknown>;
	const proto = Object.getPrototypeOf(element) as object;
	const values = new Map<string, unknown>();
	for (const name of names) {
		values.set(name, target[name]);
		delete target[name];
		const get = (): unknown => Reflect.get(proto, name, element);
		accessors.add(get);
		Object.defineProperty(element, name, {
			configurable: true,
			get,
			set(value: unknown) {
				if (!upgrading.has(element)) {
					values.delete(name);
					removeAccessor(element, name);
				}
				Reflect.set(proto, name, value, element);
			},
		});
	}
	const attributes = new Map([...element.attributes].map((a) => [a.name, a.value] as const));
	held.set(element, { values, attributes });
	queueMicrotask(() => releaseEarlyProperties(element));
}

/** Gives the held values to the class's setters. A setter that refuses its value is reported. */
export function releaseEarlyProperties(element: HTMLElement): void {
	const entry = held.get(element);
	if (!entry) return;
	held.delete(element);
	for (const name of entry.values.keys()) removeAccessor(element, name);
	const target = element as unknown as Record<string, unknown>;
	for (const [name, value] of entry.values) {
		try {
			target[name] = value;
		} catch (error) {
			queueMicrotask(() => {
				throw error;
			});
		}
	}
}

const adopting = new WeakSet<object>();

type Lifecycle = {
	connectedCallback?: () => void;
	attributeChangedCallback?: (name: string, old: string | null, value: string | null) => void;
};

/** Wraps the lifecycle callbacks of `ctor`'s elements so early properties follow the rule above. */
export function adoptEarlyProperties(ctor: CustomElementConstructor): void {
	const proto = ctor.prototype as Lifecycle;
	if (adopting.has(proto)) return;
	adopting.add(proto);
	const { connectedCallback: connected, attributeChangedCallback: changed } = proto;
	Object.defineProperty(proto, 'connectedCallback', {
		configurable: true,
		writable: true,
		value(this: HTMLElement): void {
			holdEarlyProperties(this);
			releaseEarlyProperties(this);
			connected?.call(this);
		},
	});
	if (!changed) return;
	Object.defineProperty(proto, 'attributeChangedCallback', {
		configurable: true,
		writable: true,
		value(this: HTMLElement, name: string, old: string | null, value: string | null): void {
			holdEarlyProperties(this);
			const entry = held.get(this);
			if (!entry || old !== null || entry.attributes.get(name) !== value)
				return changed.call(this, name, old, value);
			entry.attributes.delete(name);
			const outer = upgrading.has(this);
			upgrading.add(this);
			try {
				changed.call(this, name, old, value);
			} finally {
				if (!outer) upgrading.delete(this);
			}
		},
	});
}
