/**
 * The Insert > Shapes gallery as PowerPoint groups it: nine categories, each an ordered list of
 * preset geometries. Every entry is a preset the shared geometry tables can draw (a test checks
 * that), and every preset of `PRESET_SHAPE_DEFINITIONS` sits in exactly one category, so the
 * gallery offers all the shapes the renderer draws. Names are the English labels PowerPoint shows;
 * a UI translates them by `shapeGalleryNameKey` / `shapeGalleryCategoryKey`.
 */

export interface ShapeGalleryEntry {
	/** The `a:prstGeom/@prst` value (the editor folds a few UI-only tokens onto it). */
	readonly preset: string;
	/** English name, shown in the ScreenTip. */
	readonly name: string;
}

export type ShapeGalleryCategoryId =
	| 'lines'
	| 'rectangles'
	| 'basic'
	| 'blockArrows'
	| 'equation'
	| 'flowchart'
	| 'starsBanners'
	| 'callouts'
	| 'actionButtons';

export interface ShapeGalleryCategory {
	readonly id: ShapeGalleryCategoryId;
	/** English heading. */
	readonly name: string;
	readonly shapes: readonly ShapeGalleryEntry[];
}

const s = (preset: string, name: string): ShapeGalleryEntry => ({ preset, name });

export const SHAPE_GALLERY_CATEGORIES: readonly ShapeGalleryCategory[] = [
	{
		id: 'lines',
		name: 'Lines',
		shapes: [
			s('line', 'Line'),
			s('straightConnector1', 'Connector: Straight'),
			s('bentConnector3', 'Connector: Elbow'),
			s('curvedConnector3', 'Connector: Curved'),
			s('lineInv', 'Line: Inverse'),
			s('bentConnector2', 'Connector: Elbow (L)'),
			s('bentConnector4', 'Connector: Elbow (3 segments)'),
			s('bentConnector5', 'Connector: Elbow (4 segments)'),
			s('curvedConnector2', 'Connector: Curved (L)'),
			s('curvedConnector4', 'Connector: Curved (3 segments)'),
			s('curvedConnector5', 'Connector: Curved (4 segments)'),
		],
	},
	{
		id: 'rectangles',
		name: 'Rectangles',
		shapes: [
			s('rect', 'Rectangle'),
			s('roundRect', 'Rectangle: Rounded Corners'),
			s('snip1Rect', 'Rectangle: Single Corner Snipped'),
			s('snip2SameRect', 'Rectangle: Top Corners Snipped'),
			s('snip2DiagRect', 'Rectangle: Diagonal Corners Snipped'),
			s('snipRoundRect', 'Rectangle: Top Corners, One Rounded and One Snipped'),
			s('round1Rect', 'Rectangle: Single Corner Rounded'),
			s('round2SameRect', 'Rectangle: Top Corners Rounded'),
			s('round2DiagRect', 'Rectangle: Diagonal Corners Rounded'),
		],
	},
	{
		id: 'basic',
		name: 'Basic Shapes',
		shapes: [
			s('ellipse', 'Oval'),
			s('triangle', 'Isosceles Triangle'),
			s('rtTriangle', 'Right Triangle'),
			s('parallelogram', 'Parallelogram'),
			s('trapezoid', 'Trapezoid'),
			s('nonIsoscelesTrapezoid', 'Trapezoid: Non-Isosceles'),
			s('diamond', 'Diamond'),
			s('pentagon', 'Regular Pentagon'),
			s('hexagon', 'Hexagon'),
			s('heptagon', 'Heptagon'),
			s('octagon', 'Octagon'),
			s('decagon', 'Decagon'),
			s('dodecagon', 'Dodecagon'),
			s('pie', 'Partial Circle'),
			s('pieWedge', 'Pie Wedge'),
			s('chord', 'Chord'),
			s('teardrop', 'Teardrop'),
			s('frame', 'Frame'),
			s('halfFrame', 'Half Frame'),
			s('corner', 'L-Shape'),
			s('diagStripe', 'Diagonal Stripe'),
			s('plus', 'Cross'),
			s('plaque', 'Plaque'),
			s('can', 'Cylinder'),
			s('cube', 'Cube'),
			s('bevel', 'Bevel'),
			s('donut', 'Donut'),
			s('noSmoking', '"No" Symbol'),
			s('blockArc', 'Block Arc'),
			s('foldedCorner', 'Folded Corner'),
			s('smileyFace', 'Smiley Face'),
			s('heart', 'Heart'),
			s('lightningBolt', 'Lightning Bolt'),
			s('sun', 'Sun'),
			s('moon', 'Moon'),
			s('cloud', 'Cloud'),
			s('arc', 'Arc'),
			s('bracketPair', 'Double Bracket'),
			s('bracePair', 'Double Brace'),
			s('leftBracket', 'Left Bracket'),
			s('rightBracket', 'Right Bracket'),
			s('leftBrace', 'Left Brace'),
			s('rightBrace', 'Right Brace'),
			s('funnel', 'Funnel'),
			s('gear6', 'Gear: 6 Teeth'),
			s('gear9', 'Gear: 9 Teeth'),
			s('cornerTabs', 'Corner Tabs'),
			s('squareTabs', 'Square Tabs'),
			s('plaqueTabs', 'Plaque Tabs'),
			s('chartX', 'Chart: X'),
			s('chartStar', 'Chart: Star'),
			s('chartPlus', 'Chart: Plus'),
		],
	},
	{
		id: 'blockArrows',
		name: 'Block Arrows',
		shapes: [
			s('rightArrow', 'Arrow: Right'),
			s('leftArrow', 'Arrow: Left'),
			s('upArrow', 'Arrow: Up'),
			s('downArrow', 'Arrow: Down'),
			s('leftRightArrow', 'Arrow: Left-Right'),
			s('upDownArrow', 'Arrow: Up-Down'),
			s('quadArrow', 'Arrow: Quad'),
			s('leftRightUpArrow', 'Arrow: Left-Right-Up'),
			s('bentArrow', 'Arrow: Bent'),
			s('uturnArrow', 'Arrow: U-Turn'),
			s('leftUpArrow', 'Arrow: Left-Up'),
			s('bentUpArrow', 'Arrow: Bent-Up'),
			s('curvedRightArrow', 'Arrow: Curved Right'),
			s('curvedLeftArrow', 'Arrow: Curved Left'),
			s('curvedUpArrow', 'Arrow: Curved Up'),
			s('curvedDownArrow', 'Arrow: Curved Down'),
			s('stripedRightArrow', 'Arrow: Striped Right'),
			s('notchedRightArrow', 'Arrow: Notched Right'),
			s('homePlate', 'Arrow: Pentagon'),
			s('chevron', 'Arrow: Chevron'),
			s('rightArrowCallout', 'Callout: Right Arrow'),
			s('downArrowCallout', 'Callout: Down Arrow'),
			s('leftArrowCallout', 'Callout: Left Arrow'),
			s('upArrowCallout', 'Callout: Up Arrow'),
			s('leftRightArrowCallout', 'Callout: Left-Right Arrow'),
			s('upDownArrowCallout', 'Callout: Up-Down Arrow'),
			s('quadArrowCallout', 'Callout: Quad Arrow'),
			s('circularArrow', 'Arrow: Circular'),
			s('leftCircularArrow', 'Arrow: Circular Left'),
			s('leftRightCircularArrow', 'Arrow: Circular Left-Right'),
			s('swooshArrow', 'Arrow: Swoosh'),
		],
	},
	{
		id: 'equation',
		name: 'Equation Shapes',
		shapes: [
			s('mathPlus', 'Plus Sign'),
			s('mathMinus', 'Minus Sign'),
			s('mathMultiply', 'Multiplication Sign'),
			s('mathDivide', 'Division Sign'),
			s('mathEqual', 'Equal Sign'),
			s('mathNotEqual', 'Not Equal Sign'),
		],
	},
	{
		id: 'flowchart',
		name: 'Flowchart',
		shapes: [
			s('flowChartProcess', 'Flowchart: Process'),
			s('flowChartAlternateProcess', 'Flowchart: Alternate Process'),
			s('flowChartDecision', 'Flowchart: Decision'),
			s('flowChartInputOutput', 'Flowchart: Data'),
			s('flowChartPredefinedProcess', 'Flowchart: Predefined Process'),
			s('flowChartInternalStorage', 'Flowchart: Internal Storage'),
			s('flowChartDocument', 'Flowchart: Document'),
			s('flowChartMultidocument', 'Flowchart: Multidocument'),
			s('flowChartTerminator', 'Flowchart: Terminator'),
			s('flowChartPreparation', 'Flowchart: Preparation'),
			s('flowChartManualInput', 'Flowchart: Manual Input'),
			s('flowChartManualOperation', 'Flowchart: Manual Operation'),
			s('flowChartConnector', 'Flowchart: Connector'),
			s('flowChartOffpageConnector', 'Flowchart: Off-page Connector'),
			s('flowChartPunchedCard', 'Flowchart: Card'),
			s('flowChartPunchedTape', 'Flowchart: Punched Tape'),
			s('flowChartSummingJunction', 'Flowchart: Summing Junction'),
			s('flowChartOr', 'Flowchart: Or'),
			s('flowChartCollate', 'Flowchart: Collate'),
			s('flowChartSort', 'Flowchart: Sort'),
			s('flowChartExtract', 'Flowchart: Extract'),
			s('flowChartMerge', 'Flowchart: Merge'),
			s('flowChartOnlineStorage', 'Flowchart: Stored Data'),
			s('flowChartDelay', 'Flowchart: Delay'),
			s('flowChartMagneticTape', 'Flowchart: Sequential Access Storage'),
			s('flowChartMagneticDisk', 'Flowchart: Magnetic Disk'),
			s('flowChartMagneticDrum', 'Flowchart: Direct Access Storage'),
			s('flowChartDisplay', 'Flowchart: Display'),
			s('flowChartOfflineStorage', 'Flowchart: Offline Storage'),
		],
	},
	{
		id: 'starsBanners',
		name: 'Stars and Banners',
		shapes: [
			s('irregularSeal1', 'Explosion: 8 Points'),
			s('irregularSeal2', 'Explosion: 14 Points'),
			s('star4', 'Star: 4 Points'),
			s('star5', 'Star: 5 Points'),
			s('star6', 'Star: 6 Points'),
			s('star7', 'Star: 7 Points'),
			s('star8', 'Star: 8 Points'),
			s('star10', 'Star: 10 Points'),
			s('star12', 'Star: 12 Points'),
			s('star16', 'Star: 16 Points'),
			s('star24', 'Star: 24 Points'),
			s('star32', 'Star: 32 Points'),
			s('ribbon2', 'Ribbon: Tilted Up'),
			s('ribbon', 'Ribbon: Tilted Down'),
			s('ellipseRibbon2', 'Ribbon: Curved and Tilted Up'),
			s('ellipseRibbon', 'Ribbon: Curved and Tilted Down'),
			s('leftRightRibbon', 'Ribbon: Left-Right'),
			s('verticalScroll', 'Scroll: Vertical'),
			s('horizontalScroll', 'Scroll: Horizontal'),
			s('wave', 'Wave'),
			s('doubleWave', 'Double Wave'),
		],
	},
	{
		id: 'callouts',
		name: 'Callouts',
		shapes: [
			s('wedgeRectCallout', 'Speech Bubble: Rectangle'),
			s('wedgeRoundRectCallout', 'Speech Bubble: Rectangle with Corners Rounded'),
			s('wedgeEllipseCallout', 'Speech Bubble: Oval'),
			s('cloudCallout', 'Thought Bubble: Cloud'),
			s('borderCallout1', 'Callout: Line'),
			s('borderCallout2', 'Callout: Bent Line'),
			s('borderCallout3', 'Callout: Double Bent Line'),
			s('accentCallout1', 'Callout: Line with Accent Bar'),
			s('accentCallout2', 'Callout: Bent Line with Accent Bar'),
			s('accentCallout3', 'Callout: Double Bent Line with Accent Bar'),
			s('callout1', 'Callout: Line with No Border'),
			s('callout2', 'Callout: Bent Line with No Border'),
			s('callout3', 'Callout: Double Bent Line with No Border'),
			s('accentBorderCallout1', 'Callout: Line with Border and Accent Bar'),
			s('accentBorderCallout2', 'Callout: Bent Line with Border and Accent Bar'),
			s('accentBorderCallout3', 'Callout: Double Bent Line with Border and Accent Bar'),
		],
	},
	{
		id: 'actionButtons',
		name: 'Action Buttons',
		shapes: [
			s('actionButtonBackPrevious', 'Action Button: Go Back or Previous'),
			s('actionButtonForwardNext', 'Action Button: Go Forward or Next'),
			s('actionButtonBeginning', 'Action Button: Go to Beginning'),
			s('actionButtonEnd', 'Action Button: Go to End'),
			s('actionButtonHome', 'Action Button: Go Home'),
			s('actionButtonInformation', 'Action Button: Information'),
			s('actionButtonReturn', 'Action Button: Return'),
			s('actionButtonMovie', 'Action Button: Movie'),
			s('actionButtonDocument', 'Action Button: Document'),
			s('actionButtonSound', 'Action Button: Sound'),
			s('actionButtonHelp', 'Action Button: Help'),
			s('actionButtonBlank', 'Action Button: Custom'),
		],
	},
];

/** Dictionary key of a category heading. */
export const shapeGalleryCategoryKey = (id: ShapeGalleryCategoryId): string =>
	`pptx.shapes.category.${id}`;

/** Dictionary key of a shape's name. */
export const shapeGalleryNameKey = (preset: string): string => `pptx.shapes.name.${preset}`;

/** Every gallery entry, flattened, with the category it sits in. */
export const SHAPE_GALLERY_ENTRIES: readonly (ShapeGalleryEntry & {
	readonly category: ShapeGalleryCategoryId;
})[] = SHAPE_GALLERY_CATEGORIES.flatMap((category) =>
	category.shapes.map((shape) => ({ ...shape, category: category.id })),
);

/** English name of a gallery preset, or undefined for a preset the gallery does not offer. */
export function shapeGalleryName(preset: string): string | undefined {
	return SHAPE_GALLERY_ENTRIES.find((entry) => entry.preset === preset)?.name;
}
