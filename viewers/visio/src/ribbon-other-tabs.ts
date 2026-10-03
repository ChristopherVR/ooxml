import {
	check,
	command,
	commandRow,
	group,
	menu,
	stack,
	type CommandSpec,
} from './ribbon-parts.js';

const PAGES = 'Needs core page insertion.';
const MEDIA = 'Needs core image and object insertion.';
const PARTS = 'Needs core container, callout and connector edits.';
const DESIGN = 'Needs core page setup and theme edits.';
const DATA = 'Needs core data linking and data graphics.';
const PROCESS = 'Needs core diagram validation.';
const REVIEW = 'Needs core comments and proofing.';

const unsupported = (
	id: string,
	label: string,
	icon: string,
	reason: string,
	size?: CommandSpec['size'],
) => ({ id, label, icon, unsupported: reason, ...(size ? { size } : {}) }) satisfies CommandSpec;
const dropdown = (doc: Document, spec: CommandSpec & { reason: string }) =>
	menu(doc, {
		...spec,
		unsupported: spec.reason,
		items: [{ id: `${spec.id}-more`, label: `${spec.label} options`, unsupported: spec.reason }],
	});

/** Visio's Insert tab: Pages, Illustrations, Diagram Parts, Links and Text. */
export function buildInsertPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'Insert commands', [
			group(doc, 'Pages', [
				dropdown(doc, {
					id: 'blank-page',
					label: 'Blank Page',
					icon: 'visioPagesPane',
					reason: PAGES,
				}),
			]),
			group(doc, 'Illustrations', [
				command(doc, unsupported('pictures', 'Pictures', 'visioPicture', MEDIA)),
				command(
					doc,
					unsupported(
						'online-pictures',
						'Online Pictures',
						'search',
						'Documents stay local; nothing is fetched.',
					),
				),
				command(doc, unsupported('chart', 'Chart', 'visioChart', MEDIA)),
				command(doc, unsupported('cad-drawing', 'CAD Drawing', 'visioCad', MEDIA)),
			]),
			group(doc, 'Diagram Parts', [
				dropdown(doc, { id: 'container', label: 'Container', icon: 'rectangle', reason: PARTS }),
				dropdown(doc, { id: 'callout', label: 'Callout', icon: 'message', reason: PARTS }),
				dropdown(doc, {
					id: 'insert-connector',
					label: 'Connector',
					icon: 'connector',
					reason: PARTS,
				}),
			]),
			group(doc, 'Links', [
				command(doc, unsupported('link', 'Link', 'visioLink', 'Needs core hyperlink edits.')),
			]),
			group(doc, 'Text', [
				dropdown(doc, { id: 'text-box', label: 'Text Box', icon: 'textBox', reason: PARTS }),
				stack(doc, [
					command(doc, unsupported('screen-tip', 'ScreenTip', 'message', REVIEW, 'small')),
					command(
						doc,
						unsupported('field', 'Field', 'textBox', 'Needs core text field edits.', 'small'),
					),
					command(doc, unsupported('object', 'Object', 'visioPicture', MEDIA, 'small')),
				]),
				dropdown(doc, {
					id: 'symbol',
					label: 'Symbol',
					icon: 'visioSymbol',
					reason: 'Needs core text edits.',
				}),
			]),
		]),
	);
}

/** Visio's Design tab: Page Setup, Themes, Variants, Backgrounds and Layout. */
export function buildDesignPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'Design commands', [
			group(
				doc,
				'Page Setup',
				[
					dropdown(doc, {
						id: 'orientation',
						label: 'Orientation',
						icon: 'visioPagesPane',
						reason: DESIGN,
					}),
					dropdown(doc, { id: 'size', label: 'Size', icon: 'pageWidth', reason: DESIGN }),
					command(doc, unsupported('auto-size', 'Auto Size', 'fitPage', DESIGN)),
				],
				{ launcher: DESIGN },
			),
			group(doc, 'Themes', [
				dropdown(doc, { id: 'themes', label: 'Themes', icon: 'quickStyles', reason: DESIGN }),
			]),
			group(doc, 'Variants', [
				dropdown(doc, { id: 'variants', label: 'Variants', icon: 'effects', reason: DESIGN }),
			]),
			group(doc, 'Backgrounds', [
				dropdown(doc, { id: 'backgrounds', label: 'Backgrounds', icon: 'fill', reason: DESIGN }),
				dropdown(doc, {
					id: 'borders-titles',
					label: 'Borders & Titles',
					icon: 'rectangle',
					reason: DESIGN,
				}),
			]),
			group(
				doc,
				'Layout',
				[
					dropdown(doc, {
						id: 're-layout',
						label: 'Re-Layout Page',
						icon: 'position',
						reason: 'Needs core automatic layout.',
					}),
					dropdown(doc, {
						id: 'connectors',
						label: 'Connectors',
						icon: 'connector',
						reason: PARTS,
					}),
				],
				{ launcher: 'Needs core automatic layout.' },
			),
		]),
	);
}

/** Visio's Data tab. Shape Data Window maps to the viewer's shape inspector. */
export function buildDataPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'Data commands', [
			group(doc, 'External Data', [
				command(doc, unsupported('quick-import', 'Quick Import', 'visioData', DATA)),
				command(doc, unsupported('custom-import', 'Custom Import', 'visioData', DATA)),
				dropdown(doc, { id: 'refresh-all', label: 'Refresh All', icon: 'reset', reason: DATA }),
			]),
			group(doc, 'Display Data', [
				dropdown(doc, {
					id: 'data-graphics',
					label: 'Data Graphics',
					icon: 'visioChart',
					reason: DATA,
				}),
				dropdown(doc, {
					id: 'insert-legend',
					label: 'Insert Legend',
					icon: 'bullets',
					reason: DATA,
				}),
			]),
			group(doc, 'Show/Hide', [
				stack(doc, [
					check(doc, {
						id: 'shape-data-window',
						label: 'Shape Data Window',
						action: { type: 'reveal', panel: 'selection' },
					}),
					check(doc, {
						id: 'external-data-window',
						label: 'External Data Window',
						unsupported: DATA,
					}),
				]),
			]),
		]),
	);
}

/** Visio's Process tab. */
export function buildProcessPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'Process commands', [
			group(doc, 'Subprocess', [
				command(doc, unsupported('create-new', 'Create New', 'visioPagesPane', PAGES)),
				command(doc, unsupported('create-from-selection', 'Create from Selection', 'group', PAGES)),
				command(doc, unsupported('link-existing', 'Link to Existing', 'visioLink', PAGES)),
			]),
			group(doc, 'Diagram Validation', [
				dropdown(doc, {
					id: 'check-diagram',
					label: 'Check Diagram',
					icon: 'check',
					reason: PROCESS,
				}),
				stack(doc, [
					command(
						doc,
						unsupported('ignore-issue', 'Ignore This Issue', 'eyeOff', PROCESS, 'small'),
					),
					check(doc, { id: 'issues-window', label: 'Issues Window', unsupported: PROCESS }),
				]),
			]),
		]),
	);
}

/** Visio's Review tab: Proofing, Language, Comments and Reports. */
export function buildReviewPanel(doc: Document, panel: HTMLElement): void {
	panel.append(
		commandRow(doc, 'Review commands', [
			group(doc, 'Proofing', [
				command(doc, unsupported('spelling', 'Spelling', 'check', REVIEW)),
				command(doc, unsupported('thesaurus', 'Thesaurus', 'search', REVIEW)),
			]),
			group(doc, 'Language', [
				dropdown(doc, { id: 'language', label: 'Language', icon: 'message', reason: REVIEW }),
			]),
			group(doc, 'Comments', [
				command(doc, unsupported('new-comment', 'New Comment', 'message', REVIEW)),
				command(doc, unsupported('comments-pane', 'Comments Pane', 'visioInspectorPane', REVIEW)),
			]),
			group(doc, 'Reports', [
				command(
					doc,
					unsupported('shape-reports', 'Shape Reports', 'visioData', 'Needs report generation.'),
				),
			]),
		]),
	);
}
