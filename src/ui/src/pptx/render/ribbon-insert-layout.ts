/**
 * The Insert tab as PowerPoint lays it out: groups in order, each control's size, its dropdown
 * entries and its ScreenTip. Pure data plus one availability rule, so the element only draws it.
 *
 * A control with no `intent` (and no enabled menu entry) has no backing feature in this editor.
 * It is still drawn, disabled, with an honest tooltip, rather than omitted or faked.
 *
 * @module render/ribbon-insert-layout
 */
import { INSERT_COMMANDS, canRequestInsert } from './ribbon-insert-state';
import type {
	InsertCommandSpec,
	InsertLabel,
	RibbonInsertCommand,
	RibbonInsertIntent,
	RibbonInsertState,
} from './ribbon-insert-state';

export interface InsertMenuEntry {
	id: string;
	label: InsertLabel;
	/** Absent when the editor has no such feature: the entry is listed, disabled. */
	intent?: RibbonInsertIntent;
}

/** Controls whose dropdown is built from a catalogue or the host's layouts. */
export type InsertSpecialMenu = 'layouts' | 'shapes' | 'chart' | 'action';

export interface InsertControlSpec {
	/** Public customization id, `insert.<group>.<control>`. */
	control: string;
	icon: string;
	label: InsertLabel;
	/** ScreenTip description under the title. */
	tip: InsertLabel;
	/** Large (icon over label) or small (icon beside label, stacked three high). */
	large: boolean;
	/** The click action, or the split button's main action. */
	intent?: RibbonInsertIntent;
	menu?: readonly InsertMenuEntry[];
	/** Main action on top, dropdown chevron below. */
	split?: boolean;
	special?: InsertSpecialMenu;
	/** A grid drawn above the entries: the Table size picker. */
	picker?: 'table';
	/** Hidden (not just disabled) when the host reports the capability missing. */
	capability?: 'chart' | 'headerFooter' | 'field';
}

export interface InsertGroupSpec {
	id: string;
	/** A nested array is a stack of small controls. */
	items: readonly (InsertControlSpec | readonly InsertControlSpec[])[];
}

const label = (key: string, fallback: string): InsertLabel => [key, fallback];
const tip = (key: string, fallback: string): InsertLabel => [`pptx.insert.tip.${key}`, fallback];

function fromCommand(
	command: RibbonInsertCommand,
	extra: Partial<InsertControlSpec> = {},
): InsertControlSpec {
	const found = INSERT_COMMANDS.find((item) => item.command === command) as InsertCommandSpec;
	return {
		control: found.control as string,
		icon: found.icon,
		label: found.label,
		tip: found.tip,
		large: true,
		intent: { kind: 'command', value: command },
		...extra,
	};
}

const entry = (
	id: string,
	key: string,
	fallback: string,
	intent?: RibbonInsertIntent,
): InsertMenuEntry => ({ id, label: label(key, fallback), ...(intent ? { intent } : {}) });
const command = (value: RibbonInsertCommand): RibbonInsertIntent => ({ kind: 'command', value });

/** A large dropdown button that has no backing feature. */
const unsupported = (
	control: string,
	icon: string,
	key: string,
	fallback: string,
	tipFallback: string,
	extra: Partial<InsertControlSpec> = {},
): InsertControlSpec => ({
	control,
	icon,
	label: label(`pptx.insert.${key}`, fallback),
	tip: tip(key, tipFallback),
	large: true,
	...extra,
});

export const INSERT_GROUPS: readonly InsertGroupSpec[] = [
	{
		id: 'insert.slides',
		items: [
			{
				control: 'insert.slides.newSlide',
				icon: 'newSlide',
				label: label('pptx.home.newSlide', 'New Slide'),
				tip: tip('newSlide', 'Add a new slide to the presentation.'),
				large: true,
				split: true,
				intent: { kind: 'newSlide', value: null },
				special: 'layouts',
			},
			unsupported(
				'insert.slides.reuseSlides',
				'reuseSlides',
				'reuseSlides',
				'Reuse Slides',
				'Insert slides from another presentation.',
			),
		],
	},
	{
		id: 'insert.tables',
		items: [
			fromCommand('table', {
				intent: undefined,
				picker: 'table',
				menu: [
					entry(
						'insertTable',
						'pptx.insert.insertTableEllipsis',
						'Insert Table...',
						command('table'),
					),
					entry('drawTable', 'pptx.insert.drawTable', 'Draw Table'),
					entry('excelSpreadsheet', 'pptx.insert.excelSpreadsheet', 'Excel Spreadsheet'),
				],
			}),
		],
	},
	{
		id: 'insert.images',
		items: [
			fromCommand('image', {
				intent: undefined,
				menu: [
					entry('thisDevice', 'pptx.insert.thisDevice', 'This Device...', command('image')),
					entry('stockImages', 'pptx.insert.stockImages', 'Stock Images...'),
					entry('onlinePictures', 'pptx.insert.onlinePictures', 'Online Pictures...'),
				],
			}),
			unsupported(
				'insert.images.screenshot',
				'screenshot',
				'screenshot',
				'Screenshot',
				'Add a screenshot of a window or part of your screen.',
				{ menu: [entry('screenClipping', 'pptx.insert.screenClipping', 'Screen Clipping')] },
			),
			unsupported(
				'insert.images.photoAlbum',
				'photoAlbum',
				'photoAlbum',
				'Photo Album',
				'Create a presentation from a set of pictures.',
				{ menu: [entry('newPhotoAlbum', 'pptx.insert.newPhotoAlbum', 'New Photo Album...')] },
			),
		],
	},
	{
		id: 'insert.camera',
		items: [
			unsupported(
				'insert.camera.cameo',
				'camera',
				'cameo',
				'Cameo',
				'Add a live camera feed to the slide.',
			),
		],
	},
	{
		id: 'insert.illustrations',
		items: [
			{
				control: 'insert.illustrations.shapes',
				icon: 'shapes',
				label: label('pptx.drawing.shapes', 'Shapes'),
				tip: tip('shapes', 'Insert ready-made shapes such as rectangles, arrows and stars.'),
				large: true,
				special: 'shapes',
			},
			unsupported(
				'insert.illustrations.icons',
				'icons',
				'icons',
				'Icons',
				'Insert icons to communicate visually with symbols.',
			),
			unsupported(
				'insert.illustrations.threeDModels',
				'model3d',
				'models3d',
				'3D Models',
				'Insert a 3D model from a file or an online source.',
				{ menu: [entry('thisDevice3d', 'pptx.insert.thisDevice', 'This Device...')] },
			),
			fromCommand('smartArt'),
			{
				control: 'insert.illustrations.chart',
				icon: 'chart',
				label: label('pptx.ribbon.chart', 'Chart'),
				tip: tip('chart', 'Insert a chart to illustrate and compare data.'),
				large: true,
				special: 'chart',
				capability: 'chart',
			},
		],
	},
	{
		id: 'insert.forms',
		items: [
			unsupported(
				'insert.forms.forms',
				'forms',
				'forms',
				'Forms',
				'Add a form to collect responses from other people.',
			),
		],
	},
	{
		id: 'insert.addIns',
		items: [
			unsupported(
				'insert.addIns.getAddIns',
				'addIns',
				'getAddIns',
				'Get Add-ins',
				'Get Office add-ins from the store.',
			),
			unsupported(
				'insert.addIns.myAddIns',
				'addIns',
				'myAddIns',
				'My Add-ins',
				'Manage the add-ins you have installed.',
			),
		],
	},
	{
		id: 'insert.links',
		items: [
			unsupported(
				'insert.links.zoom',
				'zoomIn',
				'zoom',
				'Zoom',
				'Insert a zoom to a slide or section to navigate the presentation.',
			),
			fromCommand('link'),
			{
				control: 'insert.links.action',
				icon: 'action',
				label: label('pptx.ribbon.action', 'Action'),
				tip: tip('action', 'Add an action to the selected object, such as a button that jumps.'),
				large: true,
				special: 'action',
			},
		],
	},
	{ id: 'insert.comments', items: [fromCommand('comment')] },
	{
		id: 'insert.text',
		items: [
			fromCommand('textBox'),
			fromCommand('headerFooter', { capability: 'headerFooter' }),
			unsupported(
				'insert.text.wordArt',
				'wordArt',
				'wordArt',
				'WordArt',
				'Insert decorative text into the slide.',
			),
			[
				{
					control: 'insert.text.dateTime',
					icon: 'clock',
					label: label('pptx.insert.dateTime', 'Date & Time'),
					tip: tip('dateTime', 'Add the current date or time to the slide.'),
					large: false,
					intent: { kind: 'field', value: 'datetime' },
					capability: 'field',
				},
				{
					control: 'insert.text.slideNumber',
					icon: 'slideNumber',
					label: label('pptx.field.slideNumber', 'Slide Number'),
					tip: tip('slideNumber', 'Add the slide number to the slide.'),
					large: false,
					intent: { kind: 'field', value: 'slidenum' },
					capability: 'field',
				},
				unsupported(
					'insert.text.object',
					'object',
					'object',
					'Object',
					'Insert an embedded object, such as a document or a spreadsheet.',
					{ large: false },
				),
			],
		],
	},
	{
		id: 'insert.symbols',
		items: [
			fromCommand('equation', {
				split: true,
				menu: [
					entry(
						'newEquation',
						'pptx.insert.insertNewEquation',
						'Insert New Equation',
						command('equation'),
					),
				],
			}),
			unsupported(
				'insert.symbols.symbol',
				'symbol',
				'symbol',
				'Symbol',
				'Insert symbols that are not on your keyboard, such as copyright marks.',
				{ large: true },
			),
		],
	},
	{
		id: 'insert.media',
		items: [
			fromCommand('media', {
				intent: undefined,
				menu: [
					entry('videoThisDevice', 'pptx.insert.thisDevice', 'This Device...', command('media')),
					entry('stockVideos', 'pptx.insert.stockVideos', 'Stock Videos...'),
					entry('onlineVideos', 'pptx.insert.onlineVideos', 'Online Videos...'),
				],
			}),
			{
				control: 'insert.media.audio',
				icon: 'audio',
				label: label('pptx.insert.audio', 'Audio'),
				tip: tip('audio', 'Insert an audio file from your device.'),
				large: true,
				menu: [
					entry('audioOnMyPc', 'pptx.insert.audioOnMyPc', 'Audio on My PC...', command('media')),
					entry('recordAudio', 'pptx.insert.recordAudio', 'Record Audio...'),
				],
			},
			unsupported(
				'insert.media.screenRecording',
				'record',
				'screenRecording',
				'Screen Recording',
				'Record your screen and audio, then add the recording to the slide.',
			),
		],
	},
];

/**
 * How a narrow window shrinks the tab. Every group steps large, then small with labels, then
 * small icons only, and only then becomes a popup; groups give way in ascending `priority`
 * (ties: right to left). `lowValue` groups hold nothing the editor can run today, so they
 * become popups before any working group loses its large buttons.
 */
export const INSERT_GROUP_SCALING: Readonly<
	Record<string, { priority: number; lowValue?: boolean }>
> = {
	'insert.camera': { priority: 1, lowValue: true },
	'insert.forms': { priority: 1, lowValue: true },
	'insert.addIns': { priority: 1, lowValue: true },
	'insert.comments': { priority: 2 },
	'insert.links': { priority: 3 },
	'insert.slides': { priority: 4 },
	'insert.images': { priority: 4 },
	'insert.tables': { priority: 5 },
	'insert.media': { priority: 6 },
	'insert.symbols': { priority: 7 },
	'insert.text': { priority: 8 },
	'insert.illustrations': { priority: 9 },
};

/** What the element needs to draw one control. */
export interface InsertControlStatus {
	hidden: boolean;
	disabled: boolean;
	/** Why it is disabled: no backing feature, read-only, or nothing selected. */
	reason: 'unsupported' | 'readOnly' | 'selection' | null;
}

/** An entry is enabled when the host can act on its intent right now. */
export function insertEntryEnabled(state: RibbonInsertState, item: InsertMenuEntry): boolean {
	return item.intent !== undefined && canRequestInsert(state, item.intent);
}

export function insertControlStatus(
	state: RibbonInsertState,
	spec: InsertControlSpec,
): InsertControlStatus {
	const hidden =
		(spec.capability === 'chart' && state.chartAvailable === false) ||
		(spec.capability === 'field' && state.fieldAvailable === false) ||
		(spec.capability === 'headerFooter' && state.headerFooterAvailable === false);
	if (spec.special === 'shapes' || spec.special === 'chart' || spec.special === 'action') {
		return { hidden, disabled: !state.editable, reason: state.editable ? null : 'readOnly' };
	}
	const actions = [
		...(spec.intent ? [spec.intent] : []),
		...(spec.menu ?? []).flatMap((item) => (item.intent ? [item.intent] : [])),
		...(spec.picker ? [{ kind: 'table', rows: 1, columns: 1 } as const] : []),
		...(spec.special === 'layouts' ? [{ kind: 'newSlide', value: null } as const] : []),
	];
	if (actions.length === 0) {
		return { hidden, disabled: true, reason: 'unsupported' };
	}
	if (actions.some((intent) => canRequestInsert(state, intent))) {
		return { hidden, disabled: false, reason: null };
	}
	if (!state.editable) {
		return { hidden, disabled: true, reason: 'readOnly' };
	}
	const needsSelection = actions.some(
		(intent) => intent.kind === 'command' && intent.value === 'link',
	);
	return { hidden, disabled: true, reason: needsSelection ? 'selection' : 'unsupported' };
}
