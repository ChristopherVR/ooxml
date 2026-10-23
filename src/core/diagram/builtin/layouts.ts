// The built-in SmartArt layout definitions (`dgm:layoutDef`), shared by every product: the
// catalogue (id, title, category) and the raw XML of each definition, inflated on demand. Moved
// from `pptx/core/utils/smartart-builtin-layouts/index.ts` (3b2ecd9a6); the pptx module keeps the
// parse into its own layout-definition model and re-exports these.
import { gunzipSync, strFromU8 } from 'fflate';
import type { DiagramLayoutPreset } from '../model/data';
import { BUILTIN_SMARTART_LAYOUTS } from './catalog';
import { BUILTIN_SMARTART_LAYOUT_DATA } from './data';
import type { BuiltinSmartArtLayoutEntry } from './types';

export type { BuiltinSmartArtLayoutEntry } from './types';
export { BUILTIN_SMARTART_LAYOUTS } from './catalog';

/** The catalogue of built-in layouts (index only: no definition is inflated). */
export function listBuiltinSmartArtLayouts(): readonly BuiltinSmartArtLayoutEntry[] {
	return BUILTIN_SMARTART_LAYOUTS;
}

/** The catalogue entry for a layout `uniqueId`, or `undefined` when it is not built in. */
export function findBuiltinSmartArtLayout(id: string): BuiltinSmartArtLayoutEntry | undefined {
	return BUILTIN_SMARTART_LAYOUTS.find((entry) => entry.id === id);
}

let blob: Uint8Array | undefined;
function dataBlob(): Uint8Array {
	blob ??= Uint8Array.from(atob(BUILTIN_SMARTART_LAYOUT_DATA), (ch) => ch.charCodeAt(0));
	return blob;
}

const inflated = new Map<string, string>();

/** The `dgm:layoutDef` XML of a built-in layout; throws for an unknown id. */
export function loadBuiltinSmartArtLayoutXml(id: string): string {
	const entry = findBuiltinSmartArtLayout(id);
	if (!entry) {
		throw new Error(`Unknown built-in SmartArt layout: ${id}`);
	}
	let xml = inflated.get(id);
	if (xml === undefined) {
		xml = strFromU8(gunzipSync(dataBlob().subarray(entry.offset, entry.offset + entry.length)));
		inflated.set(id, xml);
	}
	return xml;
}

/**
 * The built-in layout a category switch applies: the layout PowerPoint offers first for that
 * family.
 */
const CATEGORY_DEFAULT_TITLE: Readonly<Record<string, string>> = {
	list: 'Basic Block List',
	process: 'Basic Process',
	cycle: 'Basic Cycle',
	hierarchy: 'Hierarchy',
	relationship: 'Basic Radial',
	matrix: 'Basic Matrix',
	pyramid: 'Basic Pyramid',
	funnel: 'Funnel',
	target: 'Basic Target',
	venn: 'Basic Venn',
	timeline: 'Basic Timeline',
	chevron: 'Basic Chevron Process',
	bending: 'Basic Bending Process',
	gear: 'Gear',
};

const idForTitle = (title: string): string | undefined =>
	BUILTIN_SMARTART_LAYOUTS.find((entry) => entry.title === title)?.id;

/** The `uniqueId` of the built-in layout for a category, or `undefined` when none is mapped. */
export function defaultBuiltinSmartArtLayoutId(category: string): string | undefined {
	const title = CATEGORY_DEFAULT_TITLE[category];
	return title === undefined ? undefined : idForTitle(title);
}

/** The built-in layout (by gallery title) each Insert SmartArt preset stands for. */
const PRESET_TITLE: Readonly<Record<DiagramLayoutPreset, string>> = {
	basicBlockList: 'Basic Block List',
	stackedList: 'Stacked List',
	horizontalBulletList: 'Horizontal Bullet List',
	tableList: 'Table List',
	basicChevronProcess: 'Basic Chevron Process',
	segmentedProcess: 'Segmented Process',
	continuousBlockProcess: 'Continuous Block Process',
	upwardArrow: 'Upward Arrow',
	basicCycle: 'Basic Cycle',
	basicRadial: 'Basic Radial',
	basicPie: 'Basic Pie',
	convergingRadial: 'Converging Radial',
	hierarchy: 'Hierarchy',
	basicVenn: 'Basic Venn',
	linearVenn: 'Linear Venn',
	alternatingHexagons: 'Alternating Hexagons',
	trapezoidList: 'Trapezoid List',
	pictureAccentList: 'Picture Accent List',
	verticalBlockList: 'Vertical Block List',
	groupedList: 'Grouped List',
	horizontalPictureList: 'Horizontal Picture List',
	verticalChevronList: 'Vertical Chevron List',
	pyramidList: 'Pyramid List',
	stepDownProcess: 'Step Down Process',
	alternatingFlow: 'Alternating Flow',
	descendingProcess: 'Descending Process',
	accentProcess: 'Accent Process',
	basicTimeline: 'Basic Timeline',
	bendingProcess: 'Basic Bending Process',
	basicTarget: 'Basic Target',
	interlockingGears: 'Gear',
	basicMatrix: 'Basic Matrix',
	basicPyramid: 'Basic Pyramid',
	invertedPyramid: 'Inverted Pyramid',
	basicFunnel: 'Funnel',
};

/** The `uniqueId` of the built-in layout an Insert SmartArt preset inserts. */
export function builtinLayoutIdForPreset(preset: DiagramLayoutPreset): string | undefined {
	return idForTitle(PRESET_TITLE[preset]);
}
