/**
 * Registration of host-workflow action groups beside the core building-block actions.
 *
 * A group is one dialog's host operations (for example Document Properties or Password
 * Protection): `id` is its key on `actions`, and `create` builds its operations from the
 * {@link BuildingBlockGroupContext} the session hands it plus the binding's own port for that
 * group. The group's type is declared once by adding its key to
 * {@link BuildingBlockActionGroupMap}, so every binding's `actions` type gains it without editing
 * the core `ViewerBuildingBlockCoreActions`. DOM-free and framework-free.
 *
 * @module render/building-blocks/action-groups
 */
import type { BuildingBlockActionId, BuildingBlockResult } from '../building-block-actions';
import type {
	BuildingBlockRequest,
	BuildingBlockRequestKind,
	BuildingBlockRequests,
} from './requests';

/**
 * The type of each registered group's operations, keyed by its `actions` key. Empty until a
 * group adds its key here (for example `documentProperties: DocumentPropertiesActions`).
 */
// oxlint-disable-next-line typescript/no-empty-object-type -- extended by each registered group
export interface BuildingBlockActionGroupMap {}

export type BuildingBlockActionGroupId = keyof BuildingBlockActionGroupMap & string;

/** What a group's operations may use from the session: gating, requests and their completion. */
export interface BuildingBlockGroupContext {
	/** The failure to return when `id` may not run now, otherwise undefined. */
	refuse(id: BuildingBlockActionId): BuildingBlockResult | undefined;
	/** Raise a host request (gated by its action's availability). */
	request(next: BuildingBlockRequest): boolean;
	/** Settle the pending request of `kind` when `result` succeeded; returns `result`. */
	complete(kind: BuildingBlockRequestKind, result: BuildingBlockResult): BuildingBlockResult;
}

/**
 * One group's definition. `Port` is what the binding supplies for it (its own state and editor
 * operations, read at call time); the same definition then serves every binding.
 */
export interface BuildingBlockActionGroup<K extends BuildingBlockActionGroupId, Port> {
	readonly id: K;
	create(context: BuildingBlockGroupContext, port: Port): BuildingBlockActionGroupMap[K];
}

/** Identity helper that checks a group definition against its declared type. */
export function defineBuildingBlockActionGroup<K extends BuildingBlockActionGroupId, Port>(
	group: BuildingBlockActionGroup<K, Port>,
): BuildingBlockActionGroup<K, Port> {
	return group;
}

/** The group context over a session's request controller and its availability. */
export function createBuildingBlockGroupContext(
	requests: BuildingBlockRequests,
	refuse: (id: BuildingBlockActionId) => BuildingBlockResult | undefined,
): BuildingBlockGroupContext {
	return {
		refuse,
		request: (next) => requests.request(next),
		complete: (kind, result) => requests.complete(kind, result),
	};
}
