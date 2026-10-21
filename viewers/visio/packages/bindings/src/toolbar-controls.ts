import type { DiagramToolbarProps } from 'ooxml-ui/visio';

/** The groups of the library toolbar; a host picks the ones it wants with `include`. */
export type ToolbarGroupId =
	| 'file'
	| 'history'
	| 'zoom'
	| 'selection'
	| 'clipboard'
	| 'tools'
	| 'search';
export const TOOLBAR_GROUPS: readonly ToolbarGroupId[] = [
	'file',
	'history',
	'zoom',
	'selection',
	'clipboard',
	'tools',
	'search',
];
export interface ToolbarButton {
	readonly kind: 'button';
	readonly id: string;
	readonly label: string;
	readonly disabled: boolean;
	/** Set for toggles such as the drawing tools. */
	readonly pressed?: boolean;
	run(): void;
}
export interface ToolbarText {
	readonly kind: 'text';
	readonly id: string;
	readonly text: string;
}
export interface ToolbarSearch {
	readonly kind: 'search';
	readonly id: 'search';
	readonly label: string;
	readonly query: string;
	/** Result summary for assistive technology, e.g. "2 of 5". */
	readonly summary: string;
	readonly previous: ToolbarButton;
	readonly next: ToolbarButton;
	change(query: string): void;
}
export type ToolbarControl = ToolbarButton | ToolbarText | ToolbarSearch;
export interface ToolbarGroup {
	readonly id: ToolbarGroupId;
	readonly controls: readonly ToolbarControl[];
}

const fire = (action: () => unknown) => () => {
	// Edit rejections surface through the controller state; a click handler never throws.
	void Promise.resolve()
		.then(action)
		.catch(() => {});
};
const button = (
	id: string,
	label: string,
	enabled: boolean,
	action: () => unknown,
	pressed?: boolean,
): ToolbarButton => ({
	kind: 'button',
	id,
	label,
	disabled: !enabled,
	...(pressed === undefined ? {} : { pressed }),
	run: fire(action),
});

/**
 * What the library Toolbar shows, as plain data every framework maps onto its own elements.
 * Every control reflects `toolbarProps`: edit actions are disabled while the diagram is read-only,
 * viewing actions (zoom, fit, search) never are, and a host dialog button appears only when the
 * host supplied its callback.
 */
export function toolbarGroups(
	props: DiagramToolbarProps,
	include: readonly ToolbarGroupId[] = TOOLBAR_GROUPS,
): ToolbarGroup[] {
	const { labels } = props;
	const has = props.pages.length > 0;
	const groups: ToolbarGroup[] = [];
	const add = (id: ToolbarGroupId, controls: ToolbarControl[]) => {
		if (include.includes(id) && controls.length) groups.push({ id, controls });
	};
	add('file', [
		...(props.onOpenFile ? [button('open', labels.Open, true, props.onOpenFile)] : []),
		...(props.onOpenPageSetup
			? [button('page-setup', labels['Page Setup'], has, props.onOpenPageSetup)]
			: []),
		...(props.onOpenShareDialog
			? [button('share', labels.Share, has, props.onOpenShareDialog)]
			: []),
		...(props.onOpenSettings
			? [button('settings', labels.Options, true, props.onOpenSettings)]
			: []),
	]);
	add('history', [
		button('undo', labels.Undo, props.canUndo, props.onUndo),
		button('redo', labels.Redo, props.canRedo, props.onRedo),
	]);
	add('zoom', [
		button('zoom-out', labels['Zoom Out'], props.canZoomOut, props.onZoomOut),
		{ kind: 'text', id: 'zoom', text: `${Math.round(props.zoom * 100)}%` },
		button('zoom-in', labels['Zoom In'], props.canZoomIn, props.onZoomIn),
		button('fit', labels['Fit to Window'], has, props.onFit),
	]);
	add('selection', [
		button('duplicate', labels.Duplicate, props.canDuplicate, props.onDuplicate),
		button('group', labels.Group, props.canGroup, props.onGroup),
		button('ungroup', labels.Ungroup, props.canUngroup, props.onUngroup),
		button('delete', labels.Delete, props.canDelete, props.onDelete),
	]);
	add('clipboard', [
		button('cut', labels.Cut, props.canCut, props.onCut),
		button('copy', labels.Copy, props.canCopy, props.onCopy),
		button('paste', labels.Paste, props.canPaste, props.onPaste),
	]);
	add(
		'tools',
		props.tools.map((tool) =>
			button(
				`tool-${tool.id}`,
				tool.label,
				tool.id === 'pointer' || props.canDraw,
				() => props.onToolChange(tool.id),
				tool.active,
			),
		),
	);
	const { search } = props;
	add('search', [
		{
			kind: 'search',
			id: 'search',
			label: labels.Find,
			query: search.query,
			summary: search.query
				? search.resultCount
					? `${search.activeIndex >= 0 ? search.activeIndex + 1 : '-'} / ${search.resultCount}`
					: '0'
				: '',
			previous: button(
				'search-previous',
				labels['Previous Result'],
				search.resultCount > 0,
				props.onPreviousSearchResult,
			),
			next: button(
				'search-next',
				labels['Next Result'],
				search.resultCount > 0,
				props.onNextSearchResult,
			),
			change: (query) => props.onSearchChange(query),
		},
	]);
	return groups;
}
