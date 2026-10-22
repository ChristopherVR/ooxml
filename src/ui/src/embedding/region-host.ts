/**
 * Mounts one host region renderer (`HostRegionRenderer` of `ooxml-core/embedding`) into one
 * container and keeps it current. Every product hosts its regions (a File tab page, a dialog
 * body, a pane) through this class, so the mount, update and teardown semantics are identical
 * in Word, Excel, PowerPoint and Visio and in every framework binding; a product only decides
 * WHERE the container goes.
 *
 * `render` remounts only when the renderer itself changes; a new context goes to `update` (a
 * mount function's handle) or the `context` property (a custom element), and only when a field
 * changed. A renderer that throws is caught and logged: it never takes the editor down.
 */
import {
	sameRegionContext,
	sameRenderer,
	type HostRegionMountResult,
	type HostRegionRenderer,
} from 'ooxml-core/embedding';

/** Attribute every region container carries, for host CSS and tests. */
export const HOST_REGION_ATTRIBUTE = 'data-office-region';

export interface HostRegionHostOptions {
	/** Context fields that are callbacks: never compared, and always call the latest one. */
	callbackKeys?: Iterable<string>;
	/** The attribute the container's region id is written to (`data-office-region`). */
	attribute?: string;
	/** Prefix of the error logged when a renderer throws. */
	label?: string;
}

/** One region container. `C` is the product's region context. */
export class HostRegionHost<C extends { readonly region: string }> {
	readonly #container: HTMLElement;
	readonly #callbackKeys: ReadonlySet<string>;
	readonly #attribute: string;
	readonly #label: string;
	#renderer: HostRegionRenderer<HTMLElement, C> | undefined;
	#context: C | undefined;
	#result: HostRegionMountResult<C> | undefined;
	#element: (HTMLElement & { context?: C }) | undefined;
	/** The latest context as the product passed it, callbacks included. */
	#latest: C | undefined;

	constructor(container: HTMLElement, options: HostRegionHostOptions = {}) {
		this.#container = container;
		this.#callbackKeys = new Set(['close', ...(options.callbackKeys ?? [])]);
		this.#attribute = options.attribute ?? HOST_REGION_ATTRIBUTE;
		this.#label = options.label ?? '[ooxml-ui]';
	}

	/** The container this host fills. */
	get container(): HTMLElement {
		return this.#container;
	}

	/** True while a renderer is mounted. */
	get mounted(): boolean {
		return this.#renderer !== undefined;
	}

	/** Mount `renderer`, or deliver a changed `context` to the mounted one. */
	render(renderer: HostRegionRenderer<HTMLElement, C>, latest: C): void {
		this.#latest = latest;
		const context = this.#stable(latest);
		this.#container.setAttribute(this.#attribute, context.region);
		if (!sameRenderer(renderer, this.#renderer)) {
			this.#unmount();
			this.#renderer = renderer;
			this.#context = context;
			this.#mount(renderer, context);
			return;
		}
		if (this.#context && sameRegionContext(this.#context, context, this.#callbackKeys)) return;
		this.#context = context;
		if (this.#element) this.#element.context = context;
		else this.#guard(() => this.#result?.update?.(context));
	}

	/** Tear the renderer down and empty the container. */
	destroy(): void {
		this.#unmount();
		this.#renderer = undefined;
		this.#context = undefined;
		this.#latest = undefined;
	}

	/**
	 * `context` with its callbacks swapped for stable ones that call the latest product
	 * callback, so a renderer that kept an older context never calls a stale closure.
	 */
	#stable(context: C): C {
		const latest = (): Record<string, unknown> =>
			(this.#latest ?? context) as unknown as Record<string, unknown>;
		const call =
			(key: string) =>
			(...args: unknown[]): unknown =>
				(latest()[key] as ((...a: unknown[]) => unknown) | undefined)?.(...args);
		const stable: Record<string, unknown> = { ...context };
		for (const key of this.#callbackKeys) if (key in context) stable[key] = call(key);
		return stable as unknown as C;
	}

	#mount(renderer: HostRegionRenderer<HTMLElement, C>, context: C): void {
		if (typeof renderer === 'function') {
			const result = this.#guard(() => renderer(this.#container, context));
			this.#result = typeof result === 'function' ? { destroy: result } : (result ?? undefined);
			return;
		}
		const spec = typeof renderer === 'string' ? { element: renderer } : renderer;
		const element = this.#guard(() => this.#container.ownerDocument.createElement(spec.element));
		if (!element) return;
		Object.assign(element, spec.properties);
		this.#element = element as HTMLElement & { context?: C };
		this.#element.context = context;
		this.#container.append(element);
	}

	#unmount(): void {
		const result = this.#result;
		this.#result = undefined;
		this.#guard(() => result?.destroy?.());
		this.#element?.remove();
		this.#element = undefined;
	}

	#guard<T>(run: () => T): T | undefined {
		try {
			return run();
		} catch (error) {
			console.error(`${this.#label} A host region renderer threw:`, error);
			return undefined;
		}
	}
}
