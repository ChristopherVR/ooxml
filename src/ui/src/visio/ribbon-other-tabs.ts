import { check, command, commandRow, group, menu, stack, type CommandSpec } from './ribbon-parts';
import { backgroundsGroup, pageSetupGroup } from './ribbon-page-setup';

const PAGES = 'Needs core page insertion.';
const OBJECTS = 'Embedded OLE objects are not supported; only pictures can be inserted.';
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
				command(doc, {
					id: 'blank-page',
					label: 'Blank Page',
					icon: 'visioPagesPane',
					action: { type: 'page-insert' },
				}),
			]),
			group(doc, 'Illustrations', [
				command(doc, {
					id: 'pictures',
					label: 'Pictures',
					icon: 'visioPicture',
					action: { type: 'insert', item: 'picture' },
				}),
				command(
					doc,
					unsupported(
						'online-pictures',
						'Online Pictures',
						'search',
						'Documents stay local; nothing is fetched.',
					),
				),
				command(
					doc,
					unsupported('chart', 'Chart', 'visioChart', 'Embedded Excel charts are not supported.'),
				),
				command(
					doc,
					unsupported(
						'cad-drawing',
						'CAD Drawing',
						'visioCad',
						'DWG and DXF import is not supported.',
					),
				),
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
				command(doc, {
					id: 'link',
					label: 'Link',
					icon: 'visioLink',
					action: { type: 'insert', item: 'link' },
					keys: ['Control+K', 'Ctrl+K'],
				}),
			]),
			group(doc, 'Text', [
				command(doc, {
					id: 'text-box',
					label: 'Text Box',
					icon: 'textBox',
					action: { type: 'tool', tool: 'text' },
				}),
				stack(doc, [
					command(doc, {
						id: 'screen-tip',
						label: 'ScreenTip',
						icon: 'message',
						size: 'small',
						action: { type: 'insert', item: 'screen-tip' },
					}),
					command(
						doc,
						unsupported(
							'field',
							'Field',
							'textBox',
							'Text fields need a Field section and <fld> text runs with evaluated formulas, which the text editor does not support.',
							'small',
						),
					),
					command(doc, unsupported('object', 'Object', 'visioPicture', OBJECTS, 'small')),
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
			pageSetupGroup(doc),
			group(doc, 'Themes', [
				dropdown(doc, { id: 'themes', label: 'Themes', icon: 'quickStyles', reason: DESIGN }),
			]),
			group(doc, 'Variants', [
				dropdown(doc, { id: 'variants', label: 'Variants', icon: 'effects', reason: DESIGN }),
			]),
			backgroundsGroup(doc),
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

/** Visio's Help tab. Help content, support and training are not bundled in this viewer. */
export function buildHelpPanel(doc: Document, panel: HTMLElement): void {
	const OFFLINE = 'Help content is not bundled; see the viewer guide.';
	panel.append(
		commandRow(doc, 'Help commands', [
			group(doc, 'Help', [
				command(doc, unsupported('help', 'Help', 'help', OFFLINE)),
				command(
					doc,
					unsupported(
						'contact-support',
						'Contact Support',
						'message',
						'No support service is connected.',
					),
				),
				command(
					doc,
					unsupported(
						'help-feedback',
						'Feedback',
						'message',
						'Feedback is not collected by this viewer.',
					),
				),
				command(doc, unsupported('show-training', 'Show Training', 'visioPresentation', OFFLINE)),
			]),
		]),
	);
}
