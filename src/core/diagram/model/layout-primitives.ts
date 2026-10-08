/**
 * Leaf DiagramML layout-definition types (localized text, categories,
 * algorithm parameters, iterator/`dgm:forEach`/`dgm:when`/`dgm:choose`
 * attributes, and `dgm:shape` metadata). `R` is the raw XML node a reader
 * keeps for round-trip (pptx: its `XmlObject` tree). Moved from
 * `pptx/core/types/smart-art-layout-primitives.ts`.
 */

export interface DiagramLocalizedText {
	value: string;
	language?: string;
}

export interface DiagramLayoutCategory {
	type: string;
	priority: number;
}

export interface DiagramAlgorithmParameter {
	type: string;
	value?: string;
}

/** Typed DiagramML CT_Algorithm data attached to a layout node. */
export interface DiagramLayoutAlgorithm {
	type: string;
	revision?: number;
	parameters?: DiagramAlgorithmParameter[];
}

export interface DiagramIteratorAttributes {
	name?: string;
	reference?: string;
	axis?: string[];
	pointTypes?: string[];
	hideLastTransition?: boolean[];
	start?: number[];
	count?: number[];
	step?: number[];
}

export interface DiagramForEach<R = unknown> extends DiagramIteratorAttributes {
	rawXml?: R;
}

export interface DiagramWhen<R = unknown> extends DiagramIteratorAttributes {
	function: string;
	argument?: string;
	operator: string;
	value: string;
	rawXml?: R;
}

export interface DiagramChoose<R = unknown> {
	name?: string;
	when: DiagramWhen<R>[];
	otherwise?: { name?: string; rawXml?: R } | null;
	rawXml?: R;
}

/** A single `dgm:adj/@val` adjustment, keyed by its `@idx` (1-based, like `a:gd`). */
export interface DiagramShapeAdjustment {
	index: number;
	value: number;
}

/**
 * Typed DiagramML CT_Shape data (`dgm:shape`) attached to a layout node: the
 * per-node preset geometry override real (and third-party/custom) layout
 * definitions use so a layoutNode can be e.g. an ellipse or a chevron instead
 * of the arranger family's hardcoded default shape.
 */
export interface DiagramLayoutNodeShape {
	/** `dgm:shape/@type`: a preset geometry name (`roundRect`, `ellipse`, `chevron`, `conn`, ...). */
	presetGeometry?: string;
	/** `dgm:adjLst/dgm:adj` entries (adjustment index -> value, as authored). */
	adjustments?: DiagramShapeAdjustment[];
	/** `dgm:shape/@hideGeom`: the shape is present only to size text, never painted. */
	hideGeometry?: boolean;
	/**
	 * `dgm:shape/@lkTxEntry` (CT_Shape, boolean, default false): this node is a
	 * decorative shape that should mirror its paired content node's text
	 * rather than always rendering blank. See `smartart-layout-interpreter-
	 * pyramid.ts`'s `arrangePyramid`, the interpreter's one existing
	 * synthesized-decorative-shape call site.
	 */
	lkTxEntry?: boolean;
}
