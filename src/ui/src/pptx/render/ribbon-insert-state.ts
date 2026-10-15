import { ACTION_BUTTON_PRESETS } from './action-buttons';
import { INSERT_CHART_TYPES } from './insert-chart';
import { shapeGalleryItemForValue } from './shape-gallery';
import { SHAPE_PRESET_DEFS } from './shape-preset-catalog';

export type RibbonInsertCommand =
	| 'textBox'
	| 'table'
	| 'image'
	| 'media'
	| 'smartArt'
	| 'equation'
	| 'link'
	| 'headerFooter'
	| 'comment';

/** A slide layout the New Slide gallery offers (the host's `layoutOptions`). */
export interface InsertSlideLayout {
	path: string;
	name: string;
}

/**
 * The layout a New Slide request means: the chosen one, else the active slide's layout, else
 * the first the deck offers (PowerPoint repeats the current layout). Shared so every binding
 * resolves the face click the same way.
 */
export function resolveInsertSlideLayout(
	layouts: readonly InsertSlideLayout[],
	requested: string | null,
	currentPath?: string,
): InsertSlideLayout | undefined {
	const wanted = requested ?? currentPath;
	return layouts.find((item) => item.path === wanted) ?? layouts[0];
}

/** The Table grid's size (PowerPoint's picker is 10 columns by 8 rows). */
export const INSERT_TABLE_MAX = { rows: 8, columns: 10 } as const;

export type RibbonInsertField = 'slidenum' | 'datetime' | 'header' | 'footer';

/**
 * Controlled Insert tab state. Native viewers own the selections, every document
 * mutation, the file/shape/SmartArt/equation/hyperlink dialogs and history.
 */
export interface RibbonInsertState {
	editable: boolean;
	/** A hyperlink always attaches to something, so Link tracks the selection. */
	hasSelection: boolean;
	/** The shape preset staged in the Shape picker. */
	shapeType: string;
	/** The chart entry staged in the Chart picker. */
	chartKind: string;
	/** The armed Freeform: Shape / Curve tool. */
	activeFreeformTool?: string | null;
	/** Visible Freeform tools; empty hides the buttons (host customization). */
	freeformTools?: readonly string[];
	/** Hosts without these capabilities hide the control instead of showing a dead one. */
	chartAvailable?: boolean;
	fieldAvailable?: boolean;
	headerFooterAvailable?: boolean;
	/**
	 * Capabilities a host opts into. A control whose capability is missing stays visible but
	 * disabled with an honest tooltip, as Office does for what a document cannot do.
	 */
	newSlideAvailable?: boolean;
	commentAvailable?: boolean;
	/** The host can open the Slide Templates dialog: New Slide's menu then always ends with it. */
	slideTemplatesAvailable?: boolean;
	/** The layouts New Slide offers; empty leaves the gallery out and inserts the host default. */
	layouts?: readonly InsertSlideLayout[];
	/** The active slide's layout, marked in the New Slide gallery. */
	currentLayoutPath?: string;
	translate?: (key: string) => string;
}

export type RibbonInsertIntent =
	| { kind: 'command'; value: RibbonInsertCommand }
	| { kind: 'shapeType'; value: string }
	| { kind: 'shape'; value: string }
	| { kind: 'chartType'; value: string }
	| { kind: 'chart'; value: string }
	| { kind: 'freeform'; value: string | null }
	| { kind: 'actionButton'; value: string }
	| { kind: 'field'; value: RibbonInsertField }
	/** New Slide; the layout path, or null for the host's default (the active slide's layout). */
	| { kind: 'newSlide'; value: string | null }
	/** Slide Templates... at the end of New Slide's menu: the host opens its dialog. */
	| { kind: 'slideTemplates' }
	/** Insert Table from the grid picker: rows by columns. */
	| { kind: 'table'; rows: number; columns: number };

export type InsertLabel = [key: string, fallback: string];

export interface InsertCommandSpec {
	command: RibbonInsertCommand;
	group: string;
	/** Public customization id; every Insert command has one in the catalogue. */
	control?: string;
	icon: string;
	label: InsertLabel;
	title: InsertLabel;
	/** The ScreenTip description under the title. */
	tip: InsertLabel;
}

const spec = (value: InsertCommandSpec): InsertCommandSpec => value;

export const INSERT_COMMANDS: readonly InsertCommandSpec[] = [
	spec({
		command: 'table',
		group: 'insert.tables',
		control: 'insert.tables.table',
		icon: 'table',
		label: ['pptx.ribbon.table', 'Table'],
		title: ['pptx.insert.insertTable', 'Insert table'],
		tip: ['pptx.insert.tip.table', 'Add a table to your slide.'],
	}),
	spec({
		command: 'image',
		group: 'insert.images',
		control: 'insert.images.pictures',
		icon: 'image',
		label: ['pptx.insert.pictures', 'Pictures'],
		title: ['pptx.ribbon.insertImage', 'Insert image'],
		tip: ['pptx.insert.tip.pictures', 'Insert pictures from your device.'],
	}),
	spec({
		command: 'smartArt',
		group: 'insert.illustrations',
		control: 'insert.illustrations.smartArt',
		icon: 'layers',
		label: ['pptx.ribbon.smartArt', 'SmartArt'],
		title: ['pptx.insert.insertSmartArt', 'Insert SmartArt'],
		tip: [
			'pptx.insert.tip.smartArt',
			'Insert a SmartArt graphic to visually communicate information.',
		],
	}),
	spec({
		command: 'link',
		group: 'insert.links',
		control: 'insert.links.link',
		icon: 'link',
		label: ['pptx.insert.link', 'Link'],
		title: ['pptx.hyperlinkDialog.title', 'Hyperlink'],
		tip: [
			'pptx.insert.tip.link',
			'Create a link to a web page, a file, an email address or a place in this presentation.',
		],
	}),
	spec({
		command: 'comment',
		group: 'insert.comments',
		control: 'insert.comments.comment',
		icon: 'messagePlus',
		label: ['pptx.insert.comment', 'Comment'],
		title: ['pptx.insert.comment', 'Comment'],
		tip: ['pptx.insert.tip.comment', 'Add a comment about the slide.'],
	}),
	spec({
		command: 'textBox',
		group: 'insert.text',
		control: 'insert.text.textBox',
		icon: 'textBox',
		label: ['pptx.ribbon.textBox', 'Text Box'],
		title: ['pptx.insert.addTextBox', 'Add text box'],
		tip: ['pptx.insert.tip.textBox', 'Add a text box to the slide.'],
	}),
	spec({
		command: 'headerFooter',
		group: 'insert.text',
		control: 'insert.text.headerFooter',
		icon: 'headerFooter',
		label: ['pptx.headerFooter.title', 'Header & Footer'],
		title: ['pptx.headerFooter.title', 'Header & Footer'],
		tip: ['pptx.insert.tip.headerFooter', 'Edit the header and footer text for the slides.'],
	}),
	spec({
		command: 'equation',
		group: 'insert.symbols',
		control: 'insert.symbols.equation',
		icon: 'equation',
		label: ['pptx.ribbon.equation', 'Equation'],
		title: ['pptx.insert.insertEquation', 'Insert Equation'],
		tip: ['pptx.insert.tip.equation', 'Insert common mathematical equations or build your own.'],
	}),
	spec({
		command: 'media',
		group: 'insert.media',
		control: 'insert.media.media',
		icon: 'video',
		label: ['pptx.insert.video', 'Video'],
		title: ['pptx.ribbon.insertMedia', 'Insert audio or video'],
		tip: ['pptx.insert.tip.video', 'Insert a video or audio file from your device.'],
	}),
];

export const INSERT_FIELDS: readonly { id: RibbonInsertField; label: InsertLabel }[] = [
	{ id: 'slidenum', label: ['pptx.field.slideNumber', 'Slide Number'] },
	{ id: 'datetime', label: ['pptx.field.dateTime', 'Date/Time'] },
	{ id: 'header', label: ['pptx.field.header', 'Header'] },
	{ id: 'footer', label: ['pptx.field.footer', 'Footer'] },
];

/** The groups in PowerPoint's order, with their captions. */
export const INSERT_GROUP_CAPTIONS: Readonly<Record<string, InsertLabel>> = {
	'insert.slides': ['pptx.insert.groupSlides', 'Slides'],
	'insert.tables': ['pptx.insert.groupTables', 'Tables'],
	'insert.images': ['pptx.insert.groupImages', 'Images'],
	'insert.camera': ['pptx.insert.groupCamera', 'Camera'],
	'insert.illustrations': ['pptx.insert.groupIllustrations', 'Illustrations'],
	'insert.forms': ['pptx.insert.groupForms', 'Forms'],
	'insert.addIns': ['pptx.insert.groupAddIns', 'Add-ins'],
	'insert.links': ['pptx.insert.groupLinks', 'Links'],
	'insert.comments': ['pptx.insert.groupComments', 'Comments'],
	'insert.text': ['pptx.insert.groupText', 'Text'],
	'insert.symbols': ['pptx.insert.groupSymbols', 'Symbols'],
	'insert.media': ['pptx.insert.groupMedia', 'Media'],
};

export function insertLabel(state: RibbonInsertState, label: InsertLabel): string {
	const value = state.translate?.(label[0]);
	return value && value !== label[0] ? value : label[1];
}

/** Whether a command has a backing feature in this host right now. */
function commandAvailable(state: RibbonInsertState, command: RibbonInsertCommand): boolean {
	switch (command) {
		case 'link':
			return state.hasSelection;
		case 'comment':
			return state.editable && state.commentAvailable === true;
		default:
			return state.editable;
	}
}

/** Reject malformed programmatic intents as well as disabled pointer/keyboard picks. */
export function canRequestInsert(state: RibbonInsertState, intent: RibbonInsertIntent): boolean {
	switch (intent.kind) {
		case 'command':
			if (!INSERT_COMMANDS.some((item) => item.command === intent.value)) {
				return false;
			}
			return commandAvailable(state, intent.value);
		case 'table':
			return (
				state.editable &&
				Number.isInteger(intent.rows) &&
				Number.isInteger(intent.columns) &&
				intent.rows >= 1 &&
				intent.rows <= INSERT_TABLE_MAX.rows &&
				intent.columns >= 1 &&
				intent.columns <= INSERT_TABLE_MAX.columns
			);
		case 'slideTemplates':
			return state.editable && state.slideTemplatesAvailable === true;
		case 'newSlide':
			return (
				state.editable &&
				state.newSlideAvailable === true &&
				(intent.value === null || (state.layouts ?? []).some((item) => item.path === intent.value))
			);
		case 'shapeType':
		case 'shape':
			return (
				state.editable &&
				(SHAPE_PRESET_DEFS.some((item) => item.type === intent.value) ||
					shapeGalleryItemForValue(intent.value)?.action.kind === 'shape')
			);
		case 'chartType':
		case 'chart':
			return (
				state.editable &&
				state.chartAvailable !== false &&
				INSERT_CHART_TYPES.some((item) => item.id === intent.value)
			);
		case 'freeform':
			return (
				state.editable &&
				(intent.value === null || (state.freeformTools ?? []).includes(intent.value))
			);
		case 'actionButton':
			return (
				state.editable && ACTION_BUTTON_PRESETS.some((item) => item.shapeType === intent.value)
			);
		case 'field':
			return (
				state.editable &&
				state.fieldAvailable !== false &&
				INSERT_FIELDS.some((item) => item.id === intent.value)
			);
	}
}
