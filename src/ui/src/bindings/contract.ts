// A product's binding contract: which props a framework forwards to its element and how, which
// element events it re-emits under which handler option, and the members of its imperative handle.
// Every adapter of the product is generated from, or checked against, this one object.
import type { Equality } from './equality';
import type { HandleSpec } from './handle';

/** How `bindElement` forwards one prop to the element. */
export type ForwardPolicy =
	/** Written on every update: settings the user cannot change inside the element. */
	| 'always'
	/** Written the first time and then only when the prop changed since it was last written. */
	| 'changed'
	/**
	 * Like `changed`, but an undefined prop is skipped and not remembered, so the element keeps
	 * what the user chose (File > Open renames the document; a later re-render must not undo it).
	 */
	| 'defined'
	/** Not forwarded by the kit: the product's mount function owns it (ordering, supersession). */
	| 'manual';

export interface PropRule<E = never, V = unknown> {
	readonly forward: ForwardPolicy;
	/** Compared and written in place of an undefined prop. */
	readonly fallback?: V;
	/** Default `Object.is`. */
	readonly same?: Equality<V>;
	/** Counts as already written at mount, so an equal first value is not written. */
	readonly initial?: V;
	/** Writes the value; default `element[key] = value`. */
	readonly apply?: (element: E, value: V) => void;
}

/**
 * Two-way binding of an event to a prop (`v-model:readOnly`, Angular `[(readOnly)]`): the adapter
 * emits `value(...args)` back to the prop when it differs from the prop's current value.
 */
export interface EventModel {
	readonly prop: string;
	readonly value: (...args: never[]) => unknown;
}

export interface EventRule<O extends string = string> {
	/** The handler option the event calls (`onDocumentChange`). */
	readonly option: O;
	/** The handler's arguments; default the event's `detail` alone. */
	readonly args?: (event: CustomEvent) => readonly unknown[];
	readonly model?: EventModel;
}

/** Argument shapes shared by the products' events. */
export const eventArgs = {
	/** The unwrapped detail (the default). */
	detail: (event: CustomEvent): readonly unknown[] => [event.detail],
	/** The event itself, for cancelable events (`file-command`: call `preventDefault()`). */
	event: (event: CustomEvent): readonly unknown[] => [event],
	/** `(detail, event)`: the detail, plus the event to cancel (`teams-open-file`). */
	detailAndEvent: (event: CustomEvent): readonly unknown[] => [event.detail, event],
	/** `(readOnly, detail)` for `readonly-change`. */
	readOnly: (event: CustomEvent<{ readOnly: boolean }>): readonly unknown[] => [
		event.detail.readOnly,
		event.detail,
	],
	/** One field of the detail (`dirty-change` carries `{ dirty }`). */
	field:
		(name: string) =>
		(event: CustomEvent): readonly unknown[] => [(event.detail as Record<string, unknown>)[name]],
} as const;

/**
 * `readonly-change` with `v-model:readOnly`: the model mirrors the host's own switch
 * (`detail.hostReadOnly`), never a collaboration role lock.
 */
export const readOnlyChange: EventRule<'onReadOnlyChange'> = {
	option: 'onReadOnlyChange',
	args: eventArgs.readOnly,
	model: {
		prop: 'readOnly',
		value: (_readOnly: boolean, detail: { hostReadOnly: boolean }) => detail.hostReadOnly,
	},
};

/**
 * The contract of one product binding. `props` lists every forwarded prop in forwarding order,
 * `events` maps each element event to its handler option, `handle` lists the handle members.
 */
export interface BindingContract<
	E = never,
	P extends object = Record<string, unknown>,
	H = unknown,
	N extends string = string,
	O extends string = string,
> {
	readonly props: { readonly [K in keyof P]-?: PropRule<E, P[K]> };
	readonly events: { readonly [K in N]: EventRule<O> };
	readonly handle: HandleSpec<H>;
}

/** What every product's mount function returns: options in, teardown out. */
export interface MountedBinding<O> {
	update(options: O): void;
	destroy(): void;
}

/** Any contract, for the functions that only read its names. */
export interface AnyContract {
	readonly props: object;
	readonly events: Readonly<Record<string, EventRule>>;
	readonly handle: object;
}

export function propKeys(contract: Pick<AnyContract, 'props'>): string[] {
	return Object.keys(contract.props);
}
export function eventNames(contract: Pick<AnyContract, 'events'>): string[] {
	return Object.keys(contract.events);
}
/** The handler option names (`onDocumentChange`), in event order. */
export function optionNames(contract: Pick<AnyContract, 'events'>): string[] {
	return Object.values(contract.events).map((rule) => rule.option);
}

/** Svelte callback props are the lower-case option: `onDocumentChange` is `ondocumentchange`. */
export const svelteHandlerName = (option: string): string => option.toLowerCase();
/** Angular outputs drop the `on`: `onDocumentChange` is `documentChange`. */
export const angularOutputName = (option: string): string =>
	option.charAt(2).toLowerCase() + option.slice(3);

/**
 * Copies the contract's props out of a framework's props or component instance. `undefined`
 * values are copied too (meaning "not provided") unless `definedOnly` drops them.
 */
export function pickProps<P extends object>(
	contract: Pick<AnyContract, 'props'>,
	source: object,
	definedOnly = false,
): P {
	const from = source as Record<string, unknown>;
	const picked: Record<string, unknown> = {};
	for (const key of propKeys(contract)) {
		const value = from[key];
		if (!definedOnly || value !== undefined) picked[key] = value;
	}
	return picked as P;
}

/** Maps handlers keyed by event name onto the contract's handler options. */
export function eventOptions<O extends object>(
	contract: Pick<AnyContract, 'events'>,
	handlers: object,
): O {
	const from = handlers as Record<string, unknown>;
	const options: Record<string, unknown> = {};
	for (const [name, rule] of Object.entries(contract.events)) options[rule.option] = from[name];
	return options as O;
}

/**
 * Reads the handler options out of a framework's props, where each one is named `nameOf(option)`
 * (identity for React and Solid, `svelteHandlerName` for Svelte). `extra` adds options that are
 * not element events (`onReady`).
 */
export function handlerOptions<O extends object>(
	contract: Pick<AnyContract, 'events'>,
	source: object,
	nameOf: (option: string) => string = (option) => option,
	extra: readonly string[] = [],
): O {
	const from = source as Record<string, unknown>;
	const options: Record<string, unknown> = {};
	for (const option of [...optionNames(contract), ...extra]) options[option] = from[nameOf(option)];
	return options as O;
}

/** Props plus handler options: what a product's mount function takes. */
export function contractOptions<O extends object>(
	contract: Pick<AnyContract, 'props' | 'events'>,
	source: object,
	nameOf?: (option: string) => string,
	extra?: readonly string[],
): O {
	return {
		...pickProps(contract, source),
		...handlerOptions(contract, source, nameOf, extra),
	} as O;
}

interface Emitter {
	emit(value: unknown): void;
}

/**
 * Handler options that emit an Angular component's outputs: each option emits the output named
 * `angularOutputName(option)` with the handler's first argument. A modelled event emits only its
 * model (`readOnlyChange` with the host's switch) and only when the input differs. `overrides`
 * replaces individual options (an output that carries two values).
 */
export function angularOutputs<O extends object>(
	contract: Pick<AnyContract, 'events'>,
	component: object,
	extra: readonly string[] = [],
	overrides: Partial<O> = {},
): O {
	const self = component as Record<string, unknown>;
	const output = (name: string) => self[name] as Emitter | undefined;
	const options: Record<string, unknown> = {};
	for (const rule of Object.values(contract.events)) {
		const model = rule.model;
		options[rule.option] = model
			? (...args: never[]) => {
					const value = model.value(...args);
					if (value !== self[model.prop]) output(`${model.prop}Change`)?.emit(value);
				}
			: (value: unknown) => output(angularOutputName(rule.option))?.emit(value);
	}
	for (const option of extra)
		options[option] = (value: unknown) => output(angularOutputName(option))?.emit(value);
	return { ...options, ...overrides } as O;
}
