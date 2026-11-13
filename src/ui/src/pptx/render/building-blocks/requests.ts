/**
 * The request lifecycle of the PowerPoint building blocks: a toolbar control or a canvas gesture
 * asks for UI the building blocks never mount (a file picker, the hyperlink, SmartArt, chart or
 * equation dialog), the host renders it and finishes the workflow with an operation, or cancels.
 *
 * {@link createBuildingBlockRequests} is a headless controller (see `render/controller`): its
 * snapshot is the pending request, so every binding reads it through its one store adapter.
 * DOM-free and framework-free.
 *
 * @module render/building-blocks/requests
 */
import type { BuildingBlockActionId, BuildingBlockResult } from '../building-block-actions';
import { createControllerStore } from '../controller/controller';
import type { PptxController } from '../controller/controller';

/**
 * The payload of each request kind. A new host workflow adds its kind here (one line) and the
 * action id that gates it to {@link BUILDING_BLOCK_REQUEST_ACTIONS}; see
 * `docs/pptx-building-blocks.md` ("Adding an action group").
 */
export interface BuildingBlockRequestMap {
	image: Record<never, never>;
	media: Record<never, never>;
	hyperlink: Record<never, never>;
	smartArt: Record<never, never>;
	chart: Record<never, never>;
	/** `existingOmml` is the equation being edited (a double-click), null for a new one. */
	equation: { existingOmml: Record<string, unknown> | null };
}

export type BuildingBlockRequestKind = keyof BuildingBlockRequestMap;

/**
 * A toolbar control or canvas gesture asked for UI the building blocks never mount. The host
 * renders it and finishes the workflow with the matching operation, or calls `cancel()`.
 */
export type BuildingBlockRequest = {
	[K in BuildingBlockRequestKind]: { kind: K } & BuildingBlockRequestMap[K];
}[BuildingBlockRequestKind];

/** The availability entry that decides whether a request of each kind may be raised. */
export const BUILDING_BLOCK_REQUEST_ACTIONS: {
	[K in BuildingBlockRequestKind]: (
		request: Extract<BuildingBlockRequest, { kind: K }>,
	) => BuildingBlockActionId;
} = {
	image: () => 'insertImage',
	media: () => 'insertMedia',
	hyperlink: () => 'hyperlink',
	smartArt: () => 'insertSmartArt',
	chart: () => 'insertChart',
	equation: (request) => (request.existingOmml ? 'editEquation' : 'insertEquation'),
};

/** The action id that gates `request`. */
export function buildingBlockRequestAction(request: BuildingBlockRequest): BuildingBlockActionId {
	const resolve = BUILDING_BLOCK_REQUEST_ACTIONS[request.kind] as (
		value: BuildingBlockRequest,
	) => BuildingBlockActionId;
	return resolve(request);
}

export interface BuildingBlockRequestsDeps {
	/** Whether the action is available now (read at call time). */
	isAvailable(id: BuildingBlockActionId): boolean;
	/** The host's `onRequest`, read at call time so a changed callback is used. */
	onRequest?(request: BuildingBlockRequest): void;
}

/** The pending request as a controller, plus its transitions. */
export interface BuildingBlockRequests extends PptxController<BuildingBlockRequest | null> {
	/**
	 * Raise `next` when its action is available: it becomes pending and `onRequest` is called.
	 * Returns false (and changes nothing) when the action is not available.
	 */
	request(next: BuildingBlockRequest): boolean;
	/** Settle the pending request of `kind` when `result` succeeded; returns `result`. */
	complete(kind: BuildingBlockRequestKind, result: BuildingBlockResult): BuildingBlockResult;
	/** Abandon the pending request without changing the deck. */
	cancel(): void;
}

export function createBuildingBlockRequests(
	deps: BuildingBlockRequestsDeps,
): BuildingBlockRequests {
	const store = createControllerStore<BuildingBlockRequest | null>(null);
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		request(next) {
			if (!deps.isAvailable(buildingBlockRequestAction(next))) {
				return false;
			}
			store.set(next);
			deps.onRequest?.(next);
			return true;
		},
		complete(kind, result) {
			if (result.ok) {
				store.update((current) => (current?.kind === kind ? null : current));
			}
			return result;
		},
		cancel: () => store.set(null),
	};
}
