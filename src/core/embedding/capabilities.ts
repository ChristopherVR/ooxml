/**
 * Capabilities: what a host lets its users do in an embedded editor, one switch per capability.
 *
 * Each product supports a subset of the shared ids below (and may add its own, such as
 * PowerPoint's `present`), so a product's capability type is a narrow union and a host that
 * names an id the product does not honour gets a compile error rather than a silent no-op.
 * `false` removes every UI path to the capability (ribbon controls, Backstage pages, menu
 * entries, shortcuts); `true` or an omitted key leaves the editor's own rules in charge.
 *
 * Capabilities are UI policy, not document protection: a host that must prevent copying or
 * export should not hand the bytes to the browser.
 */

/** Capability ids every editor understands the same way. A product honours a subset. */
export const EDITOR_CAPABILITY_IDS = [
	/** Changing the document. `false` opens it read-only, whatever the edit prop says. */
	'edit',
	/** Saving a copy in another format (File > Export, Save As, Download). */
	'export',
	/** Printing (File > Print, the Print commands and Ctrl+P). */
	'print',
	/** Sharing the document with other people (File > Share, Share buttons). */
	'share',
	/** Copying content out (Copy, Cut and their shortcuts). */
	'copy',
	/** Adding and editing comments. */
	'comment',
	/** Tracked changes: turning tracking on, accepting and rejecting. */
	'review',
	/** Find and Replace. */
	'find',
	/** File > New and File > Open. */
	'open',
	/** File > Save. */
	'save',
	/** File > Options and Customize Ribbon. */
	'settings',
] as const;

/** One shared capability id. */
export type EditorCapabilityId = (typeof EDITOR_CAPABILITY_IDS)[number];

/**
 * Per-capability switches for a product whose capability ids are `Id` (a subset of
 * `EditorCapabilityId`, optionally widened with product ids).
 */
export type EditorCapabilities<Id extends string = EditorCapabilityId> = Partial<
	Record<Id, boolean>
>;

/** True unless the host switched `id` off (`false`); an omitted key is enabled. */
export function isCapabilityEnabled<Id extends string>(
	capabilities: EditorCapabilities<Id> | null | undefined,
	id: NoInfer<Id>,
): boolean {
	return capabilities?.[id] !== false;
}

/** The ids the host switched off, in the order they were given. */
export function disabledCapabilities<Id extends string>(
	capabilities: EditorCapabilities<Id> | null | undefined,
): Id[] {
	if (!capabilities) return [];
	return (Object.keys(capabilities) as Id[]).filter((id) => capabilities[id] === false);
}

/**
 * What each capability removes in one product: the product's own action ids (ribbon controls,
 * menu entries, Backstage pages, shortcuts), keyed by capability.
 */
export type CapabilityActionMap<Id extends string, Action extends string> = Readonly<
	Partial<Record<Id, readonly Action[]>>
>;

/** Every action a disabled capability removes, deduplicated, in map order. */
export function capabilityHiddenActions<Id extends string, Action extends string>(
	capabilities: EditorCapabilities<Id> | null | undefined,
	map: CapabilityActionMap<Id, Action>,
): Action[] {
	const hidden = new Set<Action>();
	for (const id of disabledCapabilities(capabilities))
		for (const action of map[id] ?? []) hidden.add(action);
	return [...hidden];
}

/**
 * The union of several hidden-action lists (the legacy `hiddenActions` prop, the ids disabled
 * capabilities remove, the user's own Customize Ribbon choices), deduplicated in first-seen order.
 */
export function mergeHiddenActions<Action extends string>(
	...lists: ReadonlyArray<readonly Action[] | null | undefined>
): Action[] {
	const merged = new Set<Action>();
	for (const list of lists) for (const action of list ?? []) merged.add(action);
	return [...merged];
}
