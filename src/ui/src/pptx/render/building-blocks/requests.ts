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
import type {
	BuildingBlockActionId,
	BuildingBlockActionState,
	BuildingBlockResult,
} from '../building-block-actions';
import type { ShowDialogActionId } from '../building-block-show-dialogs';
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
	/** Open password, Mark as Final: render your own UI from `actions.passwordProtection`. */
	passwordProtection: Record<never, never>;
	/** Render your own editor from `actions.documentProperties`. */
	documentProperties: Record<never, never>;
	/** Render your own options from `actions.fontEmbedding`. */
	fontEmbedding: Record<never, never>;
	/** Render your own list from `actions.digitalSignatures`. */
	digitalSignatures: Record<never, never>;
	setUpSlideShow: Record<never, never>;
	customShows: Record<never, never>;
	/** The shortcut help was opened (`open`) or closed; the host shows or hides its own. */
	shortcuts: { open: boolean };
}

/** Every availability id: the core workflows and the slide show dialogs. */
export type BuildingBlockAvailabilityId = BuildingBlockActionId | ShowDialogActionId;

/** What each workflow can do right now, as `actions.availability` reports it. */
export type BuildingBlockAvailability = Record<
	BuildingBlockAvailabilityId,
	BuildingBlockActionState
>;

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
	) => BuildingBlockAvailabilityId;
} = {
	image: () => 'insertImage',
	media: () => 'insertMedia',
	hyperlink: () => 'hyperlink',
	smartArt: () => 'insertSmartArt',
	chart: () => 'insertChart',
	equation: (request) => (request.existingOmml ? 'editEquation' : 'insertEquation'),
	passwordProtection: () => 'passwordProtection',
	documentProperties: () => 'documentProperties',
	fontEmbedding: () => 'fontEmbedding',
	digitalSignatures: () => 'digitalSignatures',
	setUpSlideShow: () => 'setUpSlideShow',
	customShows: () => 'customShows',
	shortcuts: () => 'shortcuts',
};

/**
 * Requests that show state as well as change it: they are raised whenever the deck is loaded
 * (also read-only or protected); applying is refused by the operations themselves.
 */
export const BUILDING_BLOCK_VIEWABLE_REQUESTS: ReadonlySet<BuildingBlockRequestKind> = new Set([
	'passwordProtection',
	'documentProperties',
	'fontEmbedding',
	'digitalSignatures',
]);

/** The action id that gates `request`. */
export function buildingBlockRequestAction(
	request: BuildingBlockRequest,
): BuildingBlockAvailabilityId {
	const resolve = BUILDING_BLOCK_REQUEST_ACTIONS[request.kind] as (
		value: BuildingBlockRequest,
	) => BuildingBlockAvailabilityId;
	return resolve(request);
}

export interface BuildingBlockRequestsDeps {
	/** The action's state now (read at call time). */
	state(id: BuildingBlockAvailabilityId): BuildingBlockActionState;
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
	/**
	 * A toggle the host mirrors changed (the shortcut help): an open request becomes pending, a
	 * close settles a pending request of its kind, and `onRequest` hears both.
	 */
	announce(next: BuildingBlockRequest, open: boolean): void;
}

export function createBuildingBlockRequests(
	deps: BuildingBlockRequestsDeps,
): BuildingBlockRequests {
	const store = createControllerStore<BuildingBlockRequest | null>(null);
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		request(next) {
			const { available, reason } = deps.state(buildingBlockRequestAction(next));
			const allowed = BUILDING_BLOCK_VIEWABLE_REQUESTS.has(next.kind)
				? reason !== 'not-ready'
				: available;
			if (!allowed) {
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
		announce(next, open) {
			if (open) {
				store.set(next);
			} else {
				store.update((current) => (current?.kind === next.kind ? null : current));
			}
			deps.onRequest?.(next);
		},
	};
}
