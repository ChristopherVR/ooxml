/**
 * Host regions: places in an editor's layout that a host fills with its own UI (a File tab
 * page, a dialog body, a docked pane). The editor keeps the place and hands it to the host as
 * an empty container; the host fills it in any framework, or none, through ONE of two forms:
 *
 *  - a mount function `(container, context) => cleanup | handle`, or
 *  - a custom element: its tag name, or `{ element, properties }`. The editor creates it,
 *    assigns `properties` once, and keeps its `context` property current.
 *
 * The contract is DOM-free: the container type is a parameter, which every product binds to
 * `HTMLElement` in `ooxml-ui`, where the shared region host mounts it.
 */

/** A custom element a region creates (the host registers it). */
export interface HostRegionElement {
	/** The tag name, for example `acme-policies`. */
	element: string;
	/** Properties assigned to the element once, before it is attached. */
	properties?: Readonly<Record<string, unknown>>;
}

/** What a mount function may return to receive context updates and teardown. */
export interface HostRegionMountResult<Context> {
	/** Called with the new context whenever it changes. */
	update?(context: Context): void;
	/** Called once when the region unmounts. */
	destroy?(): void;
}

/** Fills `container` for one region. Returns nothing, a cleanup function or a handle. */
export type HostRegionMount<Container, Context> = (
	container: Container,
	context: Context,
) => void | (() => void) | HostRegionMountResult<Context>;

/** Any renderer form a region accepts. A string is a custom element tag name. */
export type HostRegionRenderer<Container, Context> =
	| HostRegionMount<Container, Context>
	| HostRegionElement
	| string;

/** The context every product's host region receives, at least. */
export interface HostRegionContextBase {
	/** Which region this renderer fills (for a host File tab page, `backstage.host.<id>`). */
	readonly region: string;
	/** The active UI locale (BCP 47). */
	readonly locale: string;
	/** False when the document is read-only (prop, capability or protection). */
	readonly canEdit: boolean;
	/** Dismisses the region: leaves the File tab, closes a dialog. A no-op where nothing closes. */
	readonly close: () => void;
}

/** Region id a host File tab page is mounted under. */
export type HostBackstageRegionId = `backstage.host.${string}`;

/** The region id of host File tab page `id`. */
export function hostBackstageRegionId(id: string): HostBackstageRegionId {
	return `backstage.host.${id}`;
}

/**
 * True when two renderers are the same: the same function, the same tag, or element specs with
 * the same tag and the same property values (so an inline `{ element, properties }` passed
 * again on every render does not remount).
 */
export function sameRenderer<Container, Context>(
	a: HostRegionRenderer<Container, Context> | undefined,
	b: HostRegionRenderer<Container, Context> | undefined,
): boolean {
	if (a === b) return true;
	if (!a || !b || typeof a !== 'object' || typeof b !== 'object' || a.element !== b.element)
		return false;
	const left = a.properties ?? {};
	const right = b.properties ?? {};
	const keys = Object.keys(left);
	return (
		keys.length === Object.keys(right).length &&
		keys.every((key) => Object.is(left[key], right[key]))
	);
}

/**
 * Shallow equality over two region contexts. Arrays (a selection) compare by value, and the
 * fields named in `callbackKeys` are skipped: callbacks are recreated on every render and their
 * identity is not a change.
 */
export function sameRegionContext(
	a: object,
	b: object,
	callbackKeys: ReadonlySet<string> = new Set(['close']),
): boolean {
	const left = a as Record<string, unknown>;
	const right = b as Record<string, unknown>;
	for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
		if (callbackKeys.has(key)) continue;
		const x = left[key];
		const y = right[key];
		if (Array.isArray(x) && Array.isArray(y)) {
			if (!sameItems(x, y)) return false;
		} else if (!Object.is(x, y)) return false;
	}
	return true;
}

function sameItems(a: readonly unknown[], b: readonly unknown[]): boolean {
	return a === b || (a.length === b.length && a.every((item, index) => Object.is(item, b[index])));
}
