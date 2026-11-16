// Gallery metadata of the built-in SmartArt colour and quick-style definitions. The generated
// table is `catalogue-native.ts`; the definitions are in `ooxml-core/diagram/builtin`.

/** One entry of Change Colors, in Office's gallery order. */
export interface DiagramNativeColorEntry {
	/** The short id, the last segment of `uniqueId` (`accent1_2`). */
	id: string;
	/** The definition's `uniqueId`. */
	uniqueId: string;
	/** The gallery name Office reports (en-US). */
	name: string;
	/** Office's gallery section: `mainScheme`, `colorful`, `accent1` to `accent6`. */
	category: string;
	/** What a tile needs to suggest the entry: the `node1` theme fill and line colours. */
	preview: {
		fills: string[];
		/** Per fill: a shade as a negative fraction, a tint as a positive one. */
		tones?: number[];
		line: string;
		/** `dgm:fillClrLst/@meth`. */
		mode: string;
		transparent?: boolean;
	};
}

/** One entry of SmartArt Styles, in Office's gallery order. */
export interface DiagramNativeStyleEntry {
	id: string;
	uniqueId: string;
	name: string;
	/** Office's gallery section: `simple` (Best Match for Document) or `3D`. */
	category: string;
	/** The `node1` style-matrix references and 3D traits. */
	preview: {
		line: number;
		fill: number;
		effect: number;
		/** A camera other than `orthographicFront`: the diagram is a spatial 3D scene. */
		camera?: string;
		/** The nodes carry a bevel. */
		bevel?: boolean;
	};
}
