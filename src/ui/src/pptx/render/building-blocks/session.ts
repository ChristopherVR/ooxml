/**
 * The framework-neutral session behind every binding's building-block `actions`: gating by
 * availability, the request lifecycle, the insert flows with their stale-target checks, hyperlink,
 * SmartArt, chart and equation operations, Find / Replace, layout apply / reset and the registered
 * action groups. A binding supplies one {@link BuildingBlockSessionPort} over its own editor (read
 * at call time, so operations that run long after they were handed out see the latest state) and
 * calls {@link PptxBuildingBlockSession.actions} from its reactive scope for the current object.
 *
 * DOM-free apart from `File`, which host pickers return.
 *
 * @module render/building-blocks/session
 */
import type { PptxElement, PptxLayoutOption, PptxSlide, SmartArtLayout } from 'ooxml-core/pptx';

import type { CanvasSize } from '../../types';
import {
	BUILDING_BLOCK_OK,
	BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS,
	buildingBlockFailure,
	describeBuildingBlockActions,
	validateHostImageFile,
	validateHostMediaFile,
} from '../building-block-actions';
import type {
	BuildingBlockAvailabilityInput,
	BuildingBlockResult,
	HostFileLimits,
} from '../building-block-actions';
import { describeShowDialogActions } from '../building-block-show-dialogs';
import { createImageElementFromFile } from '../image-file-insertion';
import { isValidInsertChartData } from '../insert-chart';
import type { InsertChartKind, InsertChartOptions } from '../insert-chart';
import { createBuildingBlockGroupContext } from './action-groups';
import type {
	BuildingBlockActionGroup,
	BuildingBlockActionGroupId,
	BuildingBlockActionGroupMap,
	BuildingBlockGroupContext,
} from './action-groups';
import type { ViewerBuildingBlockActions } from './actions-types';
import { buildFindReplaceActions, buildLayoutActions } from './find-layout';
import type { BuildingBlockFindReplacePort } from './find-layout';
import { hyperlinkEditDataFromAction, resolveHyperlinkEditResult } from './hyperlink-edit-data';
import type { HyperlinkEditData } from './hyperlink-edit-data';
import { createBuildingBlockRequests } from './requests';
import type {
	BuildingBlockAvailability,
	BuildingBlockAvailabilityId,
	BuildingBlockRequest,
	BuildingBlockRequests,
} from './requests';

/** The deck, slide and slide position an asynchronous insert must still target when it lands. */
export interface BuildingBlockTarget {
	/** The loaded deck's identity (the binding's handler); null before a deck loads. */
	deck: unknown;
	slideId: string | undefined;
	slideIndex: number;
}

/** What a binding supplies; every member is read at call time. */
export interface BuildingBlockSessionPort {
	/** The binding's state as `describeBuildingBlockActions` reads it; the session derives the rest. */
	availabilityInput(): BuildingBlockAvailabilityInput;
	fileLimits(): HostFileLimits | undefined;
	onRequest?(request: BuildingBlockRequest): void;
	target(): BuildingBlockTarget;
	canvasSize(): CanvasSize;
	/** The first of the effective selection, as the insert handlers read it. */
	selectedElement(): PptxElement | null;
	slideCount(): number;
	activeSlide(): PptxSlide | undefined;
	/** The layouts the active slide may switch to (its master's). */
	scopedLayouts(): PptxLayoutOption[];
	/** History-tracked insertion of a prepared element on the active slide. */
	addElement(element: PptxElement): void;
	/** Insert an audio or video file; `sameTarget` is checked before the element is added. */
	insertMediaFile(file: File, sameTarget: () => boolean): Promise<BuildingBlockResult>;
	/** Apply a validated link to the selection (refuses unsafe URLs itself). */
	applyHyperlink(data: HyperlinkEditData): BuildingBlockResult;
	removeHyperlink(): BuildingBlockResult;
	insertSmartArt(layout: SmartArtLayout, items: string[]): void;
	insertChart(kind: InsertChartKind, options?: InsertChartOptions): void;
	insertEquation(omml: Record<string, unknown>): void;
	updateEquation(omml: Record<string, unknown>): void;
	findReplace(): BuildingBlockFindReplacePort;
	/** Apply a layout to the active slide; rejects when the layout cannot be applied. */
	applyLayout(path: string): Promise<void>;
	/** Decode a picked image; defaults to `createImageElementFromFile`. */
	decodeImage?(file: File, canvasSize: CanvasSize): Promise<PptxElement | null>;
	/** `actions.cancel()` ran: close what the binding mirrors (the shortcut help). */
	onCancel?(): void;
}

/** Every availability entry for a binding's state: the core workflows and the slide show dialogs. */
export function describeBuildingBlockAvailability(
	input: BuildingBlockAvailabilityInput,
): BuildingBlockAvailability {
	return { ...describeBuildingBlockActions(input), ...describeShowDialogActions(input) };
}

/** The OMML of the element's first equation run, null when it has none. */
export function selectionEquation(element: PptxElement | null): Record<string, unknown> | null {
	const segments = (element as { textSegments?: Array<{ equationXml?: unknown }> } | null)
		?.textSegments;
	const found = segments?.find((segment) => segment.equationXml)?.equationXml;
	return found ? (found as Record<string, unknown>) : null;
}

export interface PptxBuildingBlockSession {
	/** The pending request as a controller (adapt it with the binding's store adapter). */
	readonly requests: BuildingBlockRequests;
	/** Context handed to registered groups. */
	readonly context: BuildingBlockGroupContext;
	/** The current `actions` object; call it from the binding's reactive scope. */
	actions(): ViewerBuildingBlockActions;
	/** Raise a request (what the toolbar's Picture and Video controls do). */
	request(next: BuildingBlockRequest): boolean;
	/** Home > Layout (any path the handler accepts) and Reset, as the toolbar runs them. */
	applyLayout(path: string): Promise<BuildingBlockResult>;
	resetSlide(): Promise<BuildingBlockResult>;
	/** Mirror a toggle the host shows (the shortcut help opening or closing); see `announce`. */
	announce(next: BuildingBlockRequest, open: boolean): void;
	/** Add a group's operations under its `actions` key. Returns the unregister function. */
	register<K extends BuildingBlockActionGroupId, Port>(
		group: BuildingBlockActionGroup<K, Port>,
		port: Port,
	): () => void;
}

export function createPptxBuildingBlockSession(
	port: BuildingBlockSessionPort,
): PptxBuildingBlockSession {
	const availability = () => describeBuildingBlockAvailability(port.availabilityInput());
	const refuse = (id: BuildingBlockAvailabilityId): BuildingBlockResult | undefined => {
		const { available, reason } = availability()[id];
		return available ? undefined : buildingBlockFailure(reason ?? 'not-ready');
	};
	const requests = createBuildingBlockRequests({
		state: (id) => availability()[id],
		onRequest: (request) => port.onRequest?.(request),
	});
	const context = createBuildingBlockGroupContext(
		requests,
		refuse,
		() => port.availabilityInput().ready,
	);
	const groups = new Map<string, () => unknown>();
	const gated = (
		id: BuildingBlockAvailabilityId,
		kind: BuildingBlockRequest['kind'],
		run: () => void,
	): BuildingBlockResult => {
		const denied = refuse(id);
		if (denied) {
			return denied;
		}
		run();
		return requests.complete(kind, BUILDING_BLOCK_OK);
	};

	const insertImage = async (file: File): Promise<BuildingBlockResult> => {
		const denied = refuse('insertImage');
		if (denied) {
			return denied;
		}
		const checked = validateHostImageFile(file, port.fileLimits()?.imageBytes);
		if (!checked.ok) {
			return checked;
		}
		const before = port.target();
		const decode = port.decodeImage ?? createImageElementFromFile;
		const element = await decode(file, port.canvasSize());
		if (!element) {
			return buildingBlockFailure('decode-failed');
		}
		// The deck or slide may have changed while the picture decoded.
		const now = port.target();
		if (now.deck !== before.deck || now.slideId !== before.slideId) {
			return buildingBlockFailure('stale-target');
		}
		const stillDenied = refuse('insertImage');
		if (stillDenied) {
			return stillDenied;
		}
		port.addElement(element);
		return requests.complete('image', BUILDING_BLOCK_OK);
	};

	const insertMedia = async (file: File): Promise<BuildingBlockResult> => {
		const denied = refuse('insertMedia');
		if (denied) {
			return denied;
		}
		const checked = validateHostMediaFile(file, port.fileLimits()?.mediaBytes);
		if (!checked.ok) {
			return checked;
		}
		const before = port.target();
		// A video's size is probed first; the deck, slide or its position may change meanwhile.
		const sameTarget = (): boolean => {
			const now = port.target();
			return (
				now.deck === before.deck &&
				now.slideId === before.slideId &&
				now.slideIndex === before.slideIndex &&
				availability().insertMedia.available
			);
		};
		return requests.complete('media', await port.insertMediaFile(file, sameTarget));
	};

	const layoutActions = () =>
		buildLayoutActions({
			activeSlide: port.activeSlide(),
			scopedLayouts: port.scopedLayouts(),
			getScopedLayouts: () => port.scopedLayouts(),
			getActiveSlide: () => port.activeSlide(),
			refuse: (id) => refuse(id),
			applyLayout: (path) => port.applyLayout(path),
		});

	const hyperlinkApply = (data: HyperlinkEditData): BuildingBlockResult => {
		const denied = refuse('hyperlink');
		if (denied) {
			return denied;
		}
		const resolved = resolveHyperlinkEditResult(data);
		if (!resolved.url && !resolved.action) {
			return buildingBlockFailure('invalid-target');
		}
		return requests.complete('hyperlink', port.applyHyperlink(data));
	};
	const hyperlinkRemove = (): BuildingBlockResult =>
		refuse('removeHyperlink') ?? port.removeHyperlink();
	const smartArtInsert = (layout: SmartArtLayout, items: string[]) =>
		gated('insertSmartArt', 'smartArt', () => port.insertSmartArt(layout, items));
	const chartInsert = (kind: InsertChartKind, options?: InsertChartOptions) => {
		if (options?.data !== undefined && !isValidInsertChartData(options.data)) {
			return refuse('insertChart') ?? buildingBlockFailure('invalid-data');
		}
		return gated('insertChart', 'chart', () => port.insertChart(kind, options));
	};
	const equationInsert = (omml: Record<string, unknown>) =>
		gated('insertEquation', 'equation', () => port.insertEquation(omml));
	const equationUpdate = (omml: Record<string, unknown>) =>
		gated('editEquation', 'equation', () => port.updateEquation(omml));
	const cancel = () => {
		requests.cancel();
		port.onCancel?.();
	};

	return {
		requests,
		context,
		request: (next) => requests.request(next),
		announce: (next, open) => requests.announce(next, open),
		applyLayout: (path) => layoutActions().apply(path, true),
		resetSlide: () => layoutActions().reset(),
		register(group, groupPort) {
			const create = () => group.create(context, groupPort);
			groups.set(group.id, create);
			return () => {
				if (groups.get(group.id) === create) {
					groups.delete(group.id);
				}
			};
		},
		actions() {
			const selected = port.selectedElement();
			const extensions: Partial<BuildingBlockActionGroupMap> = {};
			for (const [id, create] of groups) {
				Reflect.set(extensions, id, create());
			}
			return {
				...(extensions as BuildingBlockActionGroupMap),
				availability: availability(),
				unmountedDialogActions: BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS,
				pending: requests.getSnapshot(),
				cancel,
				insertImage,
				insertMedia,
				hyperlink: {
					current: hyperlinkEditDataFromAction(selected?.actionClick),
					slideCount: port.slideCount(),
					apply: hyperlinkApply,
					remove: hyperlinkRemove,
				},
				smartArt: { insert: smartArtInsert },
				chart: { insert: chartInsert },
				equation: {
					current: selectionEquation(selected),
					insert: equationInsert,
					update: equationUpdate,
				},
				findReplace: buildFindReplaceActions(port.findReplace(), () => refuse('replace')),
				layout: layoutActions().layout,
			};
		},
	};
}
