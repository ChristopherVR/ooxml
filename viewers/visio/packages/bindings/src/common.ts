import { mountViewer, type MountedViewer } from '../../../src/binding.js';
import type { ViewerState } from '../../../src/controller.js';
import {
	eventKeys,
	propertyKeys,
	type ViewerCallbacks,
	type ViewerOptions,
	type ViewerProperties,
} from '../../../src/contract.js';

/** Framework props permit undefined; omitted values preserve the shared binding's state. */
export type ViewerProps = { [K in keyof ViewerProperties]?: ViewerProperties[K] | undefined } & {
	events?: ViewerCallbacks | undefined;
};
export type ViewerHandle = Pick<
	MountedViewer,
	| 'element'
	| 'controller'
	| 'load'
	| 'applyEdits'
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

/** One mapping for every property and event, shared by every framework adapter. */
export function viewerOptions(props: ViewerProps): ViewerOptions {
	const result: ViewerOptions = { events: props.events ?? {} };
	for (const key of propertyKeys) {
		const value = props[key];
		if (value !== undefined) Object.assign(result, { [key]: value });
	}
	return result;
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
	const requireViewer = (): MountedViewer => {
		const binding = current();
		if (!binding) throw new Error('The viewer is not mounted.');
		return binding;
	};
	return {
		get element() {
			return requireViewer().element;
		},
		get controller() {
			return requireViewer().controller;
		},
		async load(source) {
			await requireViewer().load(source);
		},
		async applyEdits(edits) {
			await requireViewer().applyEdits(edits);
		},
		async replacePlainText(pageId, shapeId, text) {
			await requireViewer().replacePlainText(pageId, shapeId, text);
		},
		async undo() {
			await requireViewer().undo();
		},
		async redo() {
			await requireViewer().redo();
		},
		cancelEdit() {
			requireViewer().cancelEdit();
		},
		exportVsdx() {
			return requireViewer().exportVsdx();
		},
		fit() {
			requireViewer().fit();
		},
		setLayerVisibility(pageId, layerId, visible) {
			requireViewer().setLayerVisibility(pageId, layerId, visible);
		},
		resetLayerVisibility(pageId) {
			requireViewer().resetLayerVisibility(pageId);
		},
		exportSvg(options) {
			return requireViewer().exportSvg(options);
		},
		createPrintSnapshot(options) {
			return requireViewer().createPrintSnapshot(options);
		},
	};
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
			[name]: (value: unknown) => {
				try {
					(props.events?.[name] as ((value: unknown) => void) | undefined)?.(value);
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
export type { ViewerEvents, VsdxSource } from '../../../src/contract.js';
export { ViewerController, type ViewerState } from '../../../src/controller.js';
export type { SvgExportOptions, SvgExportResult } from '../../../src/export-svg.js';
export type {
	CurrentPagePrintSnapshotOptions,
	PrintSnapshot,
	PrintSnapshotPage,
	PrintSnapshotOptions,
	PrintSnapshotLimits,
	PrintSnapshotUsage,
} from '../../../src/print-snapshot.js';

export type { ViewerEditState, VsdxExportResult } from '../../../src/document-history.js';

export type { VisioEdit, VisioGeometryEdit } from 'ooxml-core/visio';
