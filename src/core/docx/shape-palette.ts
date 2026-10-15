// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The shapes Insert > Shapes offers, grouped as Word groups them. Every entry is a preset geometry
// the shared geometry tables can draw; `name` is the English label the editor translates.

export interface PaletteShape {
	preset: string;
	name: string;
}
export interface PaletteGroup {
	name: string;
	shapes: readonly PaletteShape[];
}

const shape = (preset: string, name: string): PaletteShape => ({ preset, name });

export const SHAPE_PALETTE: readonly PaletteGroup[] = [
	{
		name: 'Lines',
		shapes: [
			shape('line', 'Line'),
			shape('straightConnector1', 'Straight Connector'),
			shape('bentConnector3', 'Elbow Connector'),
			shape('curvedConnector3', 'Curved Connector'),
		],
	},
	{
		name: 'Rectangles',
		shapes: [
			shape('rect', 'Rectangle'),
			shape('roundRect', 'Rounded Rectangle'),
			shape('snip1Rect', 'Snip Single Corner'),
			shape('round2SameRect', 'Round Same-Side Corners'),
		],
	},
	{
		name: 'Basic Shapes',
		shapes: [
			shape('ellipse', 'Oval'),
			shape('triangle', 'Triangle'),
			shape('rtTriangle', 'Right Triangle'),
			shape('parallelogram', 'Parallelogram'),
			shape('trapezoid', 'Trapezoid'),
			shape('diamond', 'Diamond'),
			shape('pentagon', 'Pentagon'),
			shape('hexagon', 'Hexagon'),
			shape('octagon', 'Octagon'),
			shape('plus', 'Cross'),
			shape('heart', 'Heart'),
			shape('smileyFace', 'Smiley Face'),
			shape('sun', 'Sun'),
			shape('moon', 'Moon'),
			shape('lightningBolt', 'Lightning Bolt'),
			shape('can', 'Cylinder'),
			shape('cube', 'Cube'),
		],
	},
	{
		name: 'Block Arrows',
		shapes: [
			shape('rightArrow', 'Right Arrow'),
			shape('leftArrow', 'Left Arrow'),
			shape('upArrow', 'Up Arrow'),
			shape('downArrow', 'Down Arrow'),
			shape('leftRightArrow', 'Left-Right Arrow'),
			shape('upDownArrow', 'Up-Down Arrow'),
			shape('bentArrow', 'Bent Arrow'),
			shape('uturnArrow', 'U-Turn Arrow'),
			shape('chevron', 'Chevron'),
			shape('homePlate', 'Pentagon Arrow'),
		],
	},
	{
		name: 'Flowchart',
		shapes: [
			shape('flowChartProcess', 'Flowchart: Process'),
			shape('flowChartAlternateProcess', 'Flowchart: Alternate Process'),
			shape('flowChartDecision', 'Flowchart: Decision'),
			shape('flowChartInputOutput', 'Flowchart: Data'),
			shape('flowChartPredefinedProcess', 'Flowchart: Predefined Process'),
			shape('flowChartDocument', 'Flowchart: Document'),
			shape('flowChartTerminator', 'Flowchart: Terminator'),
			shape('flowChartConnector', 'Flowchart: Connector'),
			shape('flowChartPreparation', 'Flowchart: Preparation'),
		],
	},
	{
		name: 'Callouts',
		shapes: [
			shape('wedgeRectCallout', 'Speech Bubble: Rectangle'),
			shape('wedgeRoundRectCallout', 'Speech Bubble: Rectangle with Corners Rounded'),
			shape('wedgeEllipseCallout', 'Speech Bubble: Oval'),
			shape('cloudCallout', 'Thought Bubble: Cloud'),
			shape('borderCallout1', 'Callout: Line'),
		],
	},
	{
		name: 'Stars and Banners',
		shapes: [
			shape('star4', 'Star: 4 Points'),
			shape('star5', 'Star: 5 Points'),
			shape('star6', 'Star: 6 Points'),
			shape('star8', 'Star: 8 Points'),
			shape('wave', 'Wave'),
		],
	},
];

/** Every English name in the palette, groups included, for the locale catalogues. */
export const SHAPE_PALETTE_NAMES: readonly string[] = SHAPE_PALETTE.flatMap((group) => [
	group.name,
	...group.shapes.map((item) => item.name),
]);

/** Size a shape gets when the user clicks instead of dragging: Word's 1 inch by 0.67 inch box. */
export const DEFAULT_DRAWN_SHAPE = { widthPx: 96, heightPx: 64 } as const;

/** Presets drawn from the corner the drag started in (lines and connectors have a direction). */
export const isDirectionalPreset = (preset: string): boolean =>
	/^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(preset);
