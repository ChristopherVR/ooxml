/**
 * The compact building-blocks toolbar every port without a full ribbon component renders
 * (vanilla's `mountToolbar`, and the Vue, Svelte and Angular `Toolbar` blocks): which controls it
 * has, their order, labels (existing dictionary keys) and enabled / pressed state. The binding
 * only maps each control onto a button and calls {@link runBuildingBlockToolbarControl}.
 *
 * React keeps its full ribbon `Toolbar` (`toolbarProps`); this is the common subset the other
 * ports share. DOM-free and framework-free.
 *
 * @module render/building-blocks/toolbar
 */
import type { BuildingBlockActionId, BuildingBlockActionState } from '../building-block-actions';
import type { ToolbarActionId } from '../toolbar-actions';
import type { PptxBuildingBlockSession } from './session';

export type BuildingBlockToolbarGroupId = 'history' | 'insert' | 'editing' | 'view' | 'navigation';

export type BuildingBlockToolbarControlId =
	| 'undo'
	| 'redo'
	| 'insertImage'
	| 'insertMedia'
	| 'insertChart'
	| 'insertSmartArt'
	| 'insertEquation'
	| 'hyperlink'
	| 'find'
	| 'resetLayout'
	| 'zoomOut'
	| 'zoomIn'
	| 'zoomToFit'
	| 'previousSlide'
	| 'nextSlide';

export interface BuildingBlockToolbarControl {
	id: BuildingBlockToolbarControlId;
	/** A `pptx.*` dictionary key; translate it with the binding's translator. */
	labelKey: string;
	disabled: boolean;
	/** Set for toggles (the find bar). */
	pressed?: boolean;
}

export interface BuildingBlockToolbarGroup {
	id: BuildingBlockToolbarGroupId;
	controls: BuildingBlockToolbarControl[];
}

export interface BuildingBlockToolbarInput {
	availability: Record<BuildingBlockActionId, BuildingBlockActionState>;
	canUndo: boolean;
	canRedo: boolean;
	findOpen: boolean;
	slideIndex: number;
	slideCount: number;
	/** The host's `hiddenActions`: `undo`, `redo`, `insert`, `zoom` and `navigation` apply here. */
	hiddenActions?: readonly ToolbarActionId[] | undefined;
}

const LABELS: Record<BuildingBlockToolbarControlId, string> = {
	undo: 'pptx.toolbar.undo',
	redo: 'pptx.toolbar.redo',
	insertImage: 'pptx.insert.pictures',
	insertMedia: 'pptx.ribbon.media',
	insertChart: 'pptx.ribbon.chart',
	insertSmartArt: 'pptx.ribbon.smartArt',
	insertEquation: 'pptx.ribbon.equation',
	hyperlink: 'pptx.ribbon.link',
	find: 'pptx.toolbar.findAndReplace',
	resetLayout: 'pptx.ribbon.groupReset',
	zoomOut: 'pptx.statusBar.zoomOut',
	zoomIn: 'pptx.statusBar.zoomIn',
	zoomToFit: 'pptx.statusBar.zoomToFit',
	previousSlide: 'pptx.mobileBar.previousSlide',
	nextSlide: 'pptx.mobileBar.nextSlide',
};

/** The groups that `hiddenActions` removes as a whole, and the single buttons it removes. */
const HIDDEN_BY: Partial<Record<BuildingBlockToolbarControlId, ToolbarActionId>> = {
	undo: 'undo',
	redo: 'redo',
	insertImage: 'insert',
	insertMedia: 'insert',
	insertChart: 'insert',
	insertSmartArt: 'insert',
	insertEquation: 'insert',
	hyperlink: 'insert',
	zoomOut: 'zoom',
	zoomIn: 'zoom',
	zoomToFit: 'zoom',
	previousSlide: 'navigation',
	nextSlide: 'navigation',
};

export function describeBuildingBlockToolbar(
	input: BuildingBlockToolbarInput,
): BuildingBlockToolbarGroup[] {
	const hidden = new Set(input.hiddenActions ?? []);
	const control = (
		id: BuildingBlockToolbarControlId,
		disabled: boolean,
		pressed?: boolean,
	): BuildingBlockToolbarControl | undefined => {
		const by = HIDDEN_BY[id];
		if (by !== undefined && hidden.has(by)) {
			return undefined;
		}
		const base = { id, labelKey: LABELS[id], disabled };
		return pressed === undefined ? base : { ...base, pressed };
	};
	const action = (id: BuildingBlockToolbarControlId & BuildingBlockActionId) =>
		control(id, !input.availability[id].available);
	const groups: Array<
		[BuildingBlockToolbarGroupId, Array<BuildingBlockToolbarControl | undefined>]
	> = [
		['history', [control('undo', !input.canUndo), control('redo', !input.canRedo)]],
		[
			'insert',
			[
				action('insertImage'),
				action('insertMedia'),
				action('insertChart'),
				action('insertSmartArt'),
				action('insertEquation'),
				action('hyperlink'),
			],
		],
		['editing', [control('find', false, input.findOpen), action('resetLayout')]],
		['view', [control('zoomOut', false), control('zoomIn', false), control('zoomToFit', false)]],
		[
			'navigation',
			[
				control('previousSlide', input.slideIndex <= 0),
				control('nextSlide', input.slideIndex >= input.slideCount - 1),
			],
		],
	];
	return groups
		.map(([id, controls]) => ({
			id,
			controls: controls.filter((entry): entry is BuildingBlockToolbarControl => !!entry),
		}))
		.filter((group) => group.controls.length > 0);
}

/** The binding's view operations a toolbar control may run (its own history, zoom and slides). */
export interface BuildingBlockToolbarView {
	undo(): void;
	redo(): void;
	toggleFind(): void;
	zoomOut(): void;
	zoomIn(): void;
	zoomToFit(): void;
	previousSlide(): void;
	nextSlide(): void;
}

/**
 * Run a toolbar control: insert controls raise the host request (the building blocks mount no
 * picker or dialog), Reset runs the session's layout reset, the rest are view operations.
 */
export function runBuildingBlockToolbarControl(
	id: BuildingBlockToolbarControlId,
	session: Pick<PptxBuildingBlockSession, 'request' | 'resetSlide'>,
	view: BuildingBlockToolbarView,
): void {
	switch (id) {
		case 'insertImage':
			session.request({ kind: 'image' });
			return;
		case 'insertMedia':
			session.request({ kind: 'media' });
			return;
		case 'insertChart':
			session.request({ kind: 'chart' });
			return;
		case 'insertSmartArt':
			session.request({ kind: 'smartArt' });
			return;
		case 'insertEquation':
			session.request({ kind: 'equation', existingOmml: null });
			return;
		case 'hyperlink':
			session.request({ kind: 'hyperlink' });
			return;
		case 'resetLayout':
			void session.resetSlide();
			return;
		case 'find':
			view.toggleFind();
			return;
		default:
			view[id]();
	}
}
