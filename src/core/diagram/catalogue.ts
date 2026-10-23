// The SmartArt gallery catalogue every product's Insert SmartArt dialog and SmartArt Design
// galleries share: the insertable layout presets by category, the colour schemes with the theme
// accents they fill nodes with, and the style intensities. Pure data with English names; a
// product translates the names with its own dictionary. Moved out of the PowerPoint UI
// (`src/ui/src/pptx/render/smart-art-presets.ts` and
// `render/ribbon-galleries/smartart-gallery-patch.ts`).
import type { DiagramColorScheme, DiagramLayoutType, DiagramStyleIntensity } from './types';
import type { DiagramLayoutPreset } from './model/data';

/** The categories of the Insert SmartArt dialog, in display order. */
export type DiagramInsertCategory = 'list' | 'process' | 'cycle' | 'hierarchy' | 'relationship';

export interface DiagramInsertCategoryEntry {
	id: DiagramInsertCategory;
	/** English name. */
	label: string;
	/** English description of what the category's layouts are for. */
	description: string;
}

export const DIAGRAM_INSERT_CATEGORIES: readonly DiagramInsertCategoryEntry[] = [
	{
		id: 'list',
		label: 'List',
		description: 'Show non-sequential or grouped blocks of information.',
	},
	{
		id: 'process',
		label: 'Process',
		description: 'Show a progression, a timeline or the sequential steps of a task.',
	},
	{
		id: 'cycle',
		label: 'Cycle',
		description: 'Show a continual process or parts around a centre.',
	},
	{
		id: 'hierarchy',
		label: 'Hierarchy',
		description: 'Show hierarchical relationships, such as an organisation chart.',
	},
	{
		id: 'relationship',
		label: 'Relationship',
		description: 'Show connections, overlaps, proportions or containment between ideas.',
	},
];

/** One insertable layout of the Insert SmartArt dialog. */
export interface DiagramInsertPreset {
	/** The layout the preset inserts. */
	layout: DiagramLayoutPreset;
	/** English name. */
	label: string;
	category: DiagramInsertCategory;
	/** The node texts a new diagram starts with. */
	defaultItems: string[];
	/** The family a gallery tile draws for the preset. */
	family: DiagramLayoutType;
}

type PresetRow = [
	DiagramLayoutPreset,
	string,
	DiagramInsertCategory,
	DiagramLayoutType,
	readonly string[],
];

// prettier-ignore
const PRESET_ROWS: readonly PresetRow[] = [
	['basicBlockList', 'Basic Block List', 'list', 'list', ['Item 1', 'Item 2', 'Item 3']],
	['stackedList', 'Stacked List', 'list', 'list', ['Item 1', 'Item 2', 'Item 3']],
	['horizontalBulletList', 'Horizontal Bullet List', 'list', 'list', ['Topic 1', 'Topic 2', 'Topic 3']],
	['tableList', 'Table List', 'list', 'matrix', ['Row 1', 'Row 2', 'Row 3']],
	['basicChevronProcess', 'Chevron Process', 'process', 'chevron', ['Step 1', 'Step 2', 'Step 3']],
	['segmentedProcess', 'Segmented Process', 'process', 'process', ['Phase 1', 'Phase 2', 'Phase 3']],
	['continuousBlockProcess', 'Continuous Block Process', 'process', 'process', ['Start', 'Middle', 'End']],
	['upwardArrow', 'Upward Arrow', 'process', 'process', ['Stage 1', 'Stage 2', 'Stage 3']],
	['basicCycle', 'Basic Cycle', 'cycle', 'cycle', ['Phase 1', 'Phase 2', 'Phase 3', 'Phase 4']],
	['basicRadial', 'Basic Radial', 'cycle', 'cycle', ['Center', 'Spoke 1', 'Spoke 2', 'Spoke 3']],
	['basicPie', 'Basic Pie', 'cycle', 'target', ['Segment 1', 'Segment 2', 'Segment 3']],
	['convergingRadial', 'Converging Radial', 'cycle', 'cycle', ['Goal', 'Input 1', 'Input 2', 'Input 3']],
	['hierarchy', 'Hierarchy', 'hierarchy', 'hierarchy', ['Manager', 'Lead A', 'Lead B']],
	['basicVenn', 'Basic Venn', 'relationship', 'venn', ['Set A', 'Set B', 'Set C']],
	['linearVenn', 'Linear Venn', 'relationship', 'relationship', ['Group 1', 'Group 2', 'Group 3']],
	['alternatingHexagons', 'Alternating Hexagons', 'relationship', 'process', ['Hex 1', 'Hex 2', 'Hex 3']],
	['trapezoidList', 'Trapezoid List', 'relationship', 'funnel', ['Level 1', 'Level 2', 'Level 3']],
	['pictureAccentList', 'Picture Accent List', 'list', 'list', ['Feature 1', 'Feature 2', 'Feature 3']],
	['verticalBlockList', 'Vertical Block List', 'list', 'list', ['Block 1', 'Block 2', 'Block 3']],
	['groupedList', 'Grouped List', 'list', 'matrix', ['Item A', 'Item B', 'Item C', 'Item D']],
	['horizontalPictureList', 'Horizontal Picture List', 'list', 'process', ['Photo 1', 'Photo 2', 'Photo 3']],
	['verticalChevronList', 'Vertical Chevron List', 'list', 'chevron', ['Priority 1', 'Priority 2', 'Priority 3']],
	['pyramidList', 'Pyramid List', 'list', 'pyramid', ['Level 1', 'Level 2', 'Level 3']],
	['stepDownProcess', 'Step Down Process', 'process', 'bending', ['Step 1', 'Step 2', 'Step 3']],
	['alternatingFlow', 'Alternating Flow', 'process', 'process', ['Stage 1', 'Stage 2', 'Stage 3', 'Stage 4']],
	['descendingProcess', 'Descending Process', 'process', 'process', ['Top', 'Middle', 'Bottom']],
	['accentProcess', 'Accent Process', 'process', 'process', ['Phase 1', 'Phase 2', 'Phase 3']],
	['basicTimeline', 'Basic Timeline', 'process', 'timeline', ['2021', '2022', '2023']],
	['bendingProcess', 'Bending Process', 'process', 'bending', ['Step 1', 'Step 2', 'Step 3', 'Step 4']],
	['basicTarget', 'Basic Target', 'relationship', 'target', ['Outer', 'Middle', 'Inner']],
	['interlockingGears', 'Interlocking Gears', 'relationship', 'gear', ['Gear 1', 'Gear 2', 'Gear 3']],
	['basicMatrix', 'Basic Matrix', 'relationship', 'matrix', ['Quadrant 1', 'Quadrant 2', 'Quadrant 3', 'Quadrant 4']],
	['basicPyramid', 'Basic Pyramid', 'relationship', 'pyramid', ['Top', 'Middle', 'Base']],
	['invertedPyramid', 'Inverted Pyramid', 'relationship', 'funnel', ['Wide', 'Narrower', 'Tip']],
	['basicFunnel', 'Basic Funnel', 'relationship', 'funnel', ['Input 1', 'Input 2', 'Input 3', 'Result']],
];

/** Every insertable preset, in the order the dialog lists them within a category. */
export const DIAGRAM_INSERT_PRESETS: readonly DiagramInsertPreset[] = PRESET_ROWS.map(
	([layout, label, category, family, items]) => ({
		layout,
		label,
		category,
		family,
		defaultItems: [...items],
	}),
);

/** A colour scheme of the SmartArt Design > Change Colors gallery. */
export interface DiagramColorEntry {
	id: DiagramColorScheme;
	section: 'colorful' | 'accent1' | 'accent2';
	/** PowerPoint's gallery name. */
	name: string;
	/** Theme accents the scheme fills nodes with, in cycle order. */
	accents: ReadonlyArray<'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'>;
}

export const DIAGRAM_COLOR_SECTIONS: ReadonlyArray<{
	id: DiagramColorEntry['section'];
	title: string;
}> = [
	{ id: 'colorful', title: 'Colorful' },
	{ id: 'accent1', title: 'Accent 1' },
	{ id: 'accent2', title: 'Accent 2' },
];

export const DIAGRAM_COLOR_ENTRIES: readonly DiagramColorEntry[] = [
	{
		id: 'colorful1',
		section: 'colorful',
		name: 'Colorful - Accent Colors',
		accents: ['accent1', 'accent2', 'accent3'],
	},
	{
		id: 'colorful2',
		section: 'colorful',
		name: 'Colorful Range - Accent Colors 2 to 3',
		accents: ['accent2', 'accent3', 'accent4'],
	},
	{
		id: 'colorful3',
		section: 'colorful',
		name: 'Colorful Range - Accent Colors 3 to 4',
		accents: ['accent3', 'accent4', 'accent5'],
	},
	{
		id: 'monochromatic1',
		section: 'accent1',
		name: 'Colored Fill - Accent 1',
		accents: ['accent1'],
	},
	{
		id: 'monochromatic2',
		section: 'accent2',
		name: 'Colored Fill - Accent 2',
		accents: ['accent2'],
	},
];

/** `{{from}}`/`{{to}}` of a "Colorful Range" name, or `{{accent}}` of a "Colored Fill" one. */
export function diagramColorLabelParams(
	entry: DiagramColorEntry,
): Record<string, number> | undefined {
	const match = /(\d) to (\d)$/u.exec(entry.name) ?? /Accent (\d)$/u.exec(entry.name);
	if (!match) return undefined;
	return match[2] ? { from: Number(match[1]), to: Number(match[2]) } : { accent: Number(match[1]) };
}

/** A style intensity of the SmartArt Styles gallery, named after its PowerPoint entry. */
export interface DiagramStyleEntry {
	id: DiagramStyleIntensity;
	name: string;
}

export const DIAGRAM_STYLE_ENTRIES: readonly DiagramStyleEntry[] = [
	{ id: 'flat', name: 'Simple Fill' },
	{ id: 'moderate', name: 'Moderate Effect' },
	{ id: 'intense', name: 'Intense Effect' },
];

/** PowerPoint's gallery order for the built-in layouts' categories, with English headings. */
export const DIAGRAM_NAMED_LAYOUT_CATEGORIES: ReadonlyArray<readonly [string, string]> = [
	['list', 'List'],
	['process', 'Process'],
	['cycle', 'Cycle'],
	['hierarchy', 'Hierarchy'],
	['relationship', 'Relationship'],
	['matrix', 'Matrix'],
	['pyramid', 'Pyramid'],
	['picture', 'Picture'],
	['timeline', 'Timeline'],
	['textcard', 'Text Card'],
	['meettheteam', 'Meet the Team'],
];
