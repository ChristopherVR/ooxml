// Shared foundation of every ooxml-ui element. Each element is a Lit class with three readable
// parts: reactive properties, a `render()` template and a sibling `.css` file imported as text
// (`import css from './button.css?raw'`). Properties are declared with `declare` and set in the
// constructor so class fields never shadow Lit's accessors. Events are plain bubbling, composed
// CustomEvents.
import {
	LitElement,
	unsafeCSS,
	type CSSResult,
	type PropertyDeclaration,
	type PropertyDeclarations,
} from 'lit';
import { emit, present } from './registry.js';
import { COMMON_CONTROL_CSS } from './styles.js';
import { OFFICE_TOKENS } from './tokens.js';

/**
 * Give every bare `var(--office-x)` in a stylesheet the token's default, so a control renders
 * correctly with or without the installed theme (the same fallback `tok()` writes). Sheets stay
 * plain CSS; unknown names and private `--_x` helpers are left alone.
 */
export function withTokens(css: string): string {
	return css.replace(/var\((--office-[\w-]+)\)/g, (whole, name: string) =>
		name in OFFICE_TOKENS
			? `var(${name}, ${OFFICE_TOKENS[name as keyof typeof OFFICE_TOKENS]})`
			: whole,
	);
}

/** An element's style list: the shared control basics plus its own `.css` files. */
export const controlStyles = (...sheets: string[]): CSSResult[] => [
	unsafeCSS(COMMON_CONTROL_CSS),
	...sheets.map((sheet) => unsafeCSS(withTokens(sheet))),
];

/**
 * A boolean attribute that reflects. A property value of `''` counts as present (`checked=""`),
 * so frameworks that pass attribute idioms through properties keep working.
 */
export const flag = {
	type: Boolean,
	reflect: true,
	converter: {
		fromAttribute: (value: string | null): boolean => value !== null,
		toAttribute: (value: unknown): string | null => (present(value) ? '' : null),
	},
} satisfies PropertyDeclaration;

/** A boolean attribute that is read but never written back (`icon-only`, `caret`). */
export const readFlag = {
	type: Boolean,
	converter: {
		fromAttribute: (value: string | null): boolean => value !== null && value !== 'false',
	},
} satisfies PropertyDeclaration;

/** `pressed="true|false"` as a tri-state: absent is `undefined`, `""` and anything but `false` is true. */
export const triState = {
	reflect: true,
	converter: {
		fromAttribute: (value: string | null): boolean | undefined =>
			value === null ? undefined : value !== 'false',
		toAttribute: (value: unknown): string | null => (value === undefined ? null : String(value)),
	},
} satisfies PropertyDeclaration;

/** A `value` that reads `'on'` while its attribute is absent, like a native checkbox or radio. */
export const valueOn = {
	type: String,
	converter: { fromAttribute: (value: string | null): string => value ?? 'on' },
} satisfies PropertyDeclaration;

export type OfficeProperties = PropertyDeclarations;

/**
 * Base class of the ooxml-ui elements. Products subclass these elements (they register an alias
 * tag that keeps their own names), and the earlier hand-built elements always had their shadow DOM
 * the moment the constructor returned. So, unlike a plain Lit element:
 *
 * - The shadow root exists as soon as the constructor returns, so a subclass can attach its own
 *   styles there.
 * - Reading `shadowRoot` of a connected element that has not rendered yet renders it first, so a
 *   subclass that inspects the shadow DOM before calling `super.connectedCallback()` finds it. A
 *   microtask renders a connected element that nothing else rendered, for a subclass that
 *   overrides `connectedCallback` without calling `super`.
 * - Updates are synchronous: setting a property or attribute has redrawn the shadow DOM before
 *   the call returns, whether or not the element is connected, the contract every consumer and
 *   test already relies on. An element that prefers Lit's batched updates sets
 *   `static syncUpdates = false`.
 *
 * A constructor must not assign a property a subclass may override with its own accessor (that
 * would run the subclass's setter before its fields exist); elements keep such state in a private
 * field behind a hand-written accessor instead.
 */
export class OfficeElement extends LitElement {
	static syncUpdates = true;
	private _rendering?: boolean;
	private _ready?: boolean;

	constructor() {
		super();
		(this as { renderRoot: HTMLElement | DocumentFragment }).renderRoot = this.createRenderRoot();
		this._ready = true;
		if ((this.constructor as typeof OfficeElement).syncUpdates)
			queueMicrotask(() => {
				if (this.isConnected) this.ensureRendered();
			});
	}

	/** Render now if the element has never rendered; a no-op afterwards or while rendering. */
	ensureRendered(): void {
		if (
			this._ready &&
			!this._rendering &&
			!this.hasUpdated &&
			this.isUpdatePending &&
			(this.constructor as typeof OfficeElement).syncUpdates
		)
			this.performUpdate();
	}

	/**
	 * Reflecting properties write their attribute as soon as they are set, even before the element
	 * connects (Lit alone reflects during an update). A `flag` property also stores a real
	 * boolean whatever it is given, so a framework writing `''` (the attribute idiom for
	 * "present") reads back `true`.
	 */
	static override createProperty(name: PropertyKey, options: PropertyDeclaration = {}): void {
		if (!options.reflect || options.noAccessor || options.state)
			return super.createProperty(name, options);
		const quiet: PropertyDeclaration = { ...options, reflect: false };
		super.createProperty(name, { ...quiet, noAccessor: true });
		const store = Symbol(String(name));
		const attribute =
			typeof options.attribute === 'string' ? options.attribute : String(name).toLowerCase();
		const isFlag = options.converter === flag.converter;
		Object.defineProperty(this.prototype, name, {
			configurable: true,
			enumerable: true,
			get(this: Record<symbol, unknown>) {
				return this[store];
			},
			set(this: Record<symbol, unknown> & OfficeElement, value: unknown) {
				const old = this[store];
				const next = isFlag ? present(value) : value;
				this[store] = next;
				// The first assignment is a constructor default: a constructor may not add attributes
				// (and must not clobber the ones an upgrade finds), so Lit reflects it with the first
				// update instead.
				const deferred = old === undefined;
				if (!deferred) this.reflectAttribute(attribute, next, options);
				this.requestUpdate(name, old, deferred ? options : quiet);
			},
		});
	}

	/**
	 * Attributes the element reads straight from the DOM, without a property of the same name: a
	 * product subclass reads them too, and a framework that assigns properties it finds (React 19)
	 * must write the attribute instead. Changing one redraws the element.
	 */
	static watched: readonly string[] = [];

	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, ...this.watched];
	}

	private fromAttribute: string | undefined;

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		this.fromAttribute = name;
		try {
			super.attributeChangedCallback(name, old, value);
		} finally {
			this.fromAttribute = undefined;
		}
		if ((this.constructor as typeof OfficeElement).watched.includes(name)) this.requestUpdate();
	}

	/** Write a property's value to its attribute now, unless the attribute is where it came from. */
	protected reflectAttribute(
		attribute: string,
		value: unknown,
		options: PropertyDeclaration,
	): void {
		if (this.fromAttribute === attribute) return;
		const converter = options.converter;
		const text =
			converter && typeof converter === 'object' && converter.toAttribute
				? converter.toAttribute(value, options.type)
				: options.type === Boolean
					? value
						? ''
						: null
					: value === null || value === undefined
						? null
						: String(value);
		if (this.getAttribute(attribute) === text) return;
		if (text === null || text === undefined) this.removeAttribute(attribute);
		else this.setAttribute(attribute, String(text));
	}

	/** Dispatch a bubbling, composed event with `detail`. */
	protected fire<T>(type: string, detail: T, cancelable = false): boolean {
		return emit(this, type, detail, cancelable);
	}

	private get immediate(): boolean {
		return (this.constructor as typeof OfficeElement).syncUpdates;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		if (this.immediate && !this.hasUpdated && this.isUpdatePending) this.performUpdate();
	}

	protected override performUpdate(): void | Promise<unknown> {
		this._rendering = true;
		try {
			return super.performUpdate();
		} finally {
			this._rendering = false;
		}
	}

	override requestUpdate(
		name?: PropertyKey,
		oldValue?: unknown,
		options?: PropertyDeclaration,
	): void {
		super.requestUpdate(name, oldValue, options);
		// Only an element that has rendered redraws at once: the first render waits for connection,
		// a `shadowRoot` read or the constructor's microtask, because an upgraded element is
		// already connected while its derived constructors are still setting their defaults.
		if (
			this._ready &&
			this.immediate &&
			this.hasUpdated &&
			!this._rendering &&
			this.isUpdatePending
		)
			this.performUpdate();
	}
}

/** The platform's `shadowRoot` getter, found on the prototype chain above this class. */
function inheritedShadowRoot(): (() => ShadowRoot | null) | undefined {
	for (
		let proto = Object.getPrototypeOf(OfficeElement.prototype) as object | null;
		proto;
		proto = Object.getPrototypeOf(proto) as object | null
	) {
		const getter = Object.getOwnPropertyDescriptor(proto, 'shadowRoot')?.get;
		if (getter) return getter as () => ShadowRoot | null;
	}
	return undefined;
}

/** `shadowRoot` renders the element first when it never has (see `OfficeElement`). */
Object.defineProperty(OfficeElement.prototype, 'shadowRoot', {
	configurable: true,
	get(this: OfficeElement): ShadowRoot | null {
		// Connected only: `document.createElement` forbids a constructor from adding host attributes.
		if (this.isConnected) this.ensureRendered();
		return inheritedShadowRoot()?.call(this) ?? null;
	},
});
