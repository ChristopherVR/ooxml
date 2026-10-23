import { mountViewer, type MountedViewer } from 'ooxml-ui/visio';
import type { ViewerState } from 'ooxml-ui/visio';
import { deferredHandle, pickProps, type BindingContract } from 'ooxml-ui/bindings';
import {
	eventKeys,
	propertyKeys,
	type ViewerCallbacks,
	type ViewerOptions,
	type ViewerProperties,
} from 'ooxml-ui/visio';

/** Framework props permit undefined; omitted values preserve the shared binding's state. */
export type ViewerProps = { [K in keyof ViewerProperties]?: ViewerProperties[K] | undefined } & {
	events?: ViewerCallbacks | undefined;
};
export type ViewerHandle = Pick<
	MountedViewer,
	| 'element'
	| 'controller'
	| 'load'
	| 'createBlankDrawing'
	| 'applyEdits'
	| 'selectShapes'
	| 'selectAll'
	| 'clearSelection'
	| 'duplicateSelection'
	| 'copySelection'
	| 'cutSelection'
	| 'pasteSelection'
	| 'replacePlainText'
	| 'undo'
	| 'redo'
	| 'cancelEdit'
	| 'exportVsdx'
	| 'fit'
	| 'setLayerVisibility'
	| 'resetLayerVisibility'
	| 'exportSvg'
	| 'createPrintSnapshot'
>;

/**
 * The Visio binding contract. `mountViewer` (ooxml-ui/visio) forwards properties and events itself,
 * so every property is `manual` here: `mountFrameworkViewer` sends them as revision-aware patches.
 * Events reach native outputs through the `events` callback map (`withEventEmitter`).
 */
export const VIEWER_CONTRACT: BindingContract<never, ViewerProperties, ViewerHandle> = {
	props: Object.fromEntries(
		propertyKeys.map((key) => [key, { forward: 'manual' }]),
	) as BindingContract<never, ViewerProperties>['props'],
	events: {},
	handle: {
		element: 'get',
		controller: 'get',
		load: 'async',
		createBlankDrawing: 'async',
		applyEdits: 'async',
		selectShapes: 'call',
		selectAll: 'call',
		clearSelection: 'call',
		duplicateSelection: 'async',
		copySelection: 'async',
		cutSelection: 'async',
		pasteSelection: 'async',
		replacePlainText: 'async',
		undo: 'async',
		redo: 'async',
		cancelEdit: 'call',
		exportVsdx: 'call',
		fit: 'call',
		setLayerVisibility: 'call',
		resetLayerVisibility: 'call',
		exportSvg: 'call',
		createPrintSnapshot: 'call',
	},
};
/** The error every lazy handle member throws before mount and after unmount. */
export const VIEWER_HANDLE_OPTIONS = { message: 'The viewer is not mounted.' } as const;

/** One mapping for every property and event, shared by every framework adapter. */
export function viewerOptions(props: ViewerProps): ViewerOptions {
	return { events: props.events ?? {}, ...pickProps<ViewerOptions>(VIEWER_CONTRACT, props, true) };
}

/**
 * Native props are snapshots, while the browser binding accepts explicit patches. Forward only
 * changed properties so a callback/class update cannot undo an imperative load or user zoom.
 */
export function mountFrameworkViewer(
	container: HTMLElement,
	initial: ViewerOptions,
): MountedViewer {
	const binding = mountViewer(container, initial);
	let previous: ViewerOptions = { ...initial };
	let destroyed = false;
	let updateRevision = 0;
	const activePatches = new Set<ViewerOptions>();
	return {
		...binding,
		update(next) {
			if (destroyed) throw new Error('The viewer has been destroyed.');
			const revision = ++updateRevision;
			const snapshot: ViewerOptions = { ...next };
			const patch: ViewerOptions = { events: snapshot.events ?? {} };
			for (const key of propertyKeys) {
				const touchedByOlderCall = [...activePatches].some((active) => key in active);
				if (snapshot[key] !== undefined && (snapshot[key] !== previous[key] || touchedByOlderCall))
					Object.assign(patch, { [key]: snapshot[key] });
			}
			// A host that stops passing readOnly is no longer restricting: reset it, as false would.
			if (snapshot.readOnly === undefined && previous.readOnly !== undefined)
				Object.assign(patch, { readOnly: false });
			// Nested updates compare against committed inputs, never an older patch's pending values.
			// They must also restore unchanged inputs touched by an older call still in progress.
			// A failed or superseded call leaves the latest successful snapshot intact.
			activePatches.add(patch);
			try {
				binding.update(patch);
			} finally {
				activePatches.delete(patch);
			}
			if (!destroyed && revision === updateRevision) previous = snapshot;
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			previous = {};
			binding.destroy();
		},
	};
}

/** Lazy imperative methods reject both before mount and after unmount. */
export function viewerHandle(current: () => MountedViewer | undefined): ViewerHandle {
	return deferredHandle(VIEWER_CONTRACT.handle, current, VIEWER_HANDLE_OPTIONS);
}

/**
 * Framework-neutral view of the controller's state for native hooks, composables, signals and
 * stores. `current` returns the mounted handle or nothing; before mount and after unmount the
 * snapshot is `null` and subscribing is a no-op that the binding repeats once mounted.
 */
export interface ViewerStateSource {
	subscribe(onChange: () => void): () => void;
	getSnapshot(): ViewerState | null;
}
export function viewerStateSource(
	current: () => Pick<ViewerHandle, 'controller'> | null | undefined,
): ViewerStateSource {
	const controller = () => {
		try {
			return current()?.controller;
		} catch {
			// Lazy handles throw before mount and after unmount; that is "no state yet".
			return undefined;
		}
	};
	return {
		subscribe(onChange) {
			return controller()?.subscribe(() => onChange()) ?? (() => {});
		},
		getSnapshot() {
			return controller()?.state ?? null;
		},
	};
}

/** Native framework emits and callback-map props share one complete event inventory. */
export function withEventEmitter(
	props: ViewerProps,
	emit: (name: keyof ViewerCallbacks, value: unknown) => void,
): ViewerOptions {
	const options = viewerOptions(props);
	const callbacks: ViewerCallbacks = {};
	for (const name of eventKeys) {
		Object.assign(callbacks, {
			// The host's own callback decides a cancelable event: `false` cancels it.
			[name]: (value: unknown) => {
				try {
					return (props.events?.[name] as ((value: unknown) => unknown) | undefined)?.(value);
				} finally {
					emit(name, value);
				}
			},
		});
	}
	return { ...options, events: callbacks };
}
export { mountViewer };
export type { MountedViewer, ViewerCallbacks, ViewerOptions, ViewerProperties };
export { eventKeys, propertyKeys };
export type { ViewerEvents, VsdxSource } from 'ooxml-ui/visio';
export type { VisioShapeSelection } from 'ooxml-ui/visio';
export { ViewerController, type ViewerState } from 'ooxml-ui/visio';
export type { SvgExportOptions, SvgExportResult } from 'ooxml-ui/visio';
export type {
	CurrentPagePrintSnapshotOptions,
	PrintSnapshot,
	PrintSnapshotPage,
	PrintSnapshotOptions,
	PrintSnapshotLimits,
	PrintSnapshotUsage,
} from 'ooxml-ui/visio';

export type { ViewerEditState, VsdxExportResult } from 'ooxml-ui/visio';
export type {
	OfficeRibbonAddInEvent,
	RibbonAddInCommand,
	RibbonAddInGroup,
	RibbonAddInTab,
} from 'ooxml-ui/visio';

export type { VisioEdit, VisioGeometryEdit, CreateVsdxOptions } from 'ooxml-core/visio';
