/**
 * The assistant handle every binding's `/ai` entry returns from `createAiAssistant(config)`,
 * and the registry factory behind it.
 *
 * A handle is opaque. Its type, {@link PptxAiAssistant}, is a brand with no readable members,
 * and its value is an empty frozen object. What it stands for (the host's config and the
 * binding's chat panel loader) sits in a registry that {@link createAiAssistantRegistry}
 * creates. This module keeps no registry of its own and exports no way to add to one: each
 * binding creates its instance in a module private to the binding (not in its package
 * exports), which its main entry and its `/ai` entry share. A value a host builds itself,
 * whatever its shape, and a handle from a registry the host creates are therefore never read
 * as a handle by a viewer.
 *
 * Each binding's viewer reads its `ai` prop, input or option through its registry's
 * `resolve`: a handle that registry created gives its config and panel loader; any other
 * value is a plain config, passed on unvalidated as before (a config the session cannot use
 * ends in the panel's error state).
 *
 * @module render/controller/ai-assistant
 */

/**
 * An AI assistant for a viewer's `ai` prop, input or option. Create it with
 * `createAiAssistant(config)` from the binding's `/ai` entry; it has no members to read or
 * construct.
 */
export interface PptxAiAssistant {
	/**
	 * A type brand only: no handle has this property at runtime. It is a string key rather
	 * than a `unique symbol`, so every declaration of it is the same type: two installs of
	 * `ooxml-ui`, or a binding entry that inlines its declarations, still accept each other's
	 * handles at compile time.
	 */
	readonly 'ooxml-ui/pptx/ai-assistant': never;
}

/** The `ai` value a viewer works with, after {@link AiAssistantRegistry.resolve}. */
export interface ResolvedAiAssistant<TConfig, TPanel> {
	/** The handle's config, or the plain config the host passed. */
	readonly config: TConfig;
	/** The handle's panel; `undefined` for a plain config (the viewer uses its own). */
	readonly panel: TPanel | undefined;
}

export interface ResolveAiAssistantOptions {
	/**
	 * Whether to warn about a handle this registry cannot read. Defaults to
	 * `process.env.NODE_ENV !== 'production'`, which a host bundler replaces; without a
	 * bundler (no `process`), it is off. Angular passes `isDevMode()`.
	 */
	readonly development?: boolean;
	/** Where the warning goes. Defaults to `console.warn`. */
	readonly warn?: (message: string) => void;
}

export interface AiAssistantRegistryOptions {
	/** The package whose `/ai` entry creates the handles (`pptx-react-viewer`), for messages. */
	readonly owner: string;
}

/** One binding's handles. The binding keeps the instance private to its package. */
export interface AiAssistantRegistry<TConfig, TPanel> {
	/**
	 * Create a handle for `config` and the binding's `panel`. Only the binding's
	 * `createAiAssistant` calls this; creating a handle loads nothing.
	 */
	register(config: TConfig, panel: TPanel): PptxAiAssistant;
	/**
	 * Read a viewer's `ai` value. A handle this registry created gives its config and panel.
	 * `undefined` and `null` give `undefined`. Any other value is a plain config with no
	 * panel, including a handle from another registry (another binding, or another copy of
	 * the package), about which this warns once in development.
	 */
	resolve(
		value: TConfig | PptxAiAssistant | null | undefined,
		options?: ResolveAiAssistantOptions,
	): ResolvedAiAssistant<TConfig, TPanel> | undefined;
}

/**
 * The tag every handle carries on its prototype. Every registry, in any copy of this module,
 * tags its handles the same way, so a registry can tell a handle it cannot read from a plain
 * config and warn about it.
 */
const HANDLE_TAG = 'PptxAiAssistant';
const handlePrototype: object = Object.freeze(
	Object.create(null, { [Symbol.toStringTag]: { value: HANDLE_TAG } }),
);
const warned = new WeakSet<object>();

/** Create one binding's registry. Each binding calls this once, in a module of its own. */
export function createAiAssistantRegistry<TConfig, TPanel>({
	owner,
}: AiAssistantRegistryOptions): AiAssistantRegistry<TConfig, TPanel> {
	const assistants = new WeakMap<object, ResolvedAiAssistant<TConfig, TPanel>>();
	return Object.freeze({
		register(config: TConfig, panel: TPanel): PptxAiAssistant {
			const handle = Object.freeze(Object.create(handlePrototype)) as PptxAiAssistant;
			assistants.set(handle, Object.freeze({ config, panel }));
			return handle;
		},
		resolve(
			value: TConfig | PptxAiAssistant | null | undefined,
			options: ResolveAiAssistantOptions = {},
		): ResolvedAiAssistant<TConfig, TPanel> | undefined {
			if (value === undefined || value === null) {
				return undefined;
			}
			// WeakMap#get returns undefined for a key that is not an object.
			const entry = assistants.get(value as object);
			if (entry) {
				return entry;
			}
			if (isHandleShaped(value) && (options.development ?? isDevelopment())) {
				warnOnce(value, unreadableHandleMessage(owner), options.warn);
			}
			return { config: value as TConfig, panel: undefined };
		},
	});
}

/** Tagged like a handle (the tag sits on the prototype, so it survives a second copy). */
function isHandleShaped(value: unknown): value is object {
	return (
		typeof value === 'object' &&
		value !== null &&
		Object.prototype.toString.call(value) === `[object ${HANDLE_TAG}]`
	);
}

function unreadableHandleMessage(owner: string): string {
	return (
		`[pptx-viewer] The \`ai\` value is a createAiAssistant handle that ${owner}/ai did not ` +
		'create: it comes from another binding, or from another copy of the package (a second ' +
		`install, or the ESM and CommonJS builds mixed). ${owner} treats it as a plain config, ` +
		`and the assistant shows its error state. Create the handle with createAiAssistant from ` +
		`${owner}/ai, loaded from the same copy of the package as the viewer.`
	);
}

function warnOnce(
	value: object,
	message: string,
	warn: ((message: string) => void) | undefined,
): void {
	if (warned.has(value)) {
		return;
	}
	warned.add(value);
	(warn ?? console.warn)(message);
}

// Module-scoped, so it neither needs nor clashes with Node's global typings.
declare const process: { readonly env: { readonly NODE_ENV?: string } };

/** `process.env.NODE_ENV !== 'production'`; false when there is no `process`. */
function isDevelopment(): boolean {
	try {
		return process.env.NODE_ENV !== 'production';
	} catch {
		return false;
	}
}
