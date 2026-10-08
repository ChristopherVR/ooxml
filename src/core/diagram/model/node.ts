/**
 * SmartArt data-model node types: per-run text, per-node visual override, and
 * the node itself. `S` is the reader's resolved run style (pptx: its
 * `TextStyle`); the layout engine never reads it. Moved from
 * `pptx/core/types/smart-art-node.ts`.
 */

import type { DiagramNodeCustomLayout } from '../types';

/**
 * A single run of text inside a SmartArt node, capturing the run text and the
 * raw `a:rPr` run-properties object verbatim so per-run formatting (bold,
 * colour, size, etc.) survives a load -> edit -> save round-trip instead of
 * collapsing to a single unstyled run.
 *
 * @example
 * ```ts
 * const run: DiagramNodeTextRun = {
 *   text: "Bold",
 *   rPr: { "@_b": "1", "@_lang": "en-US" },
 * };
 * // => satisfies DiagramNodeTextRun
 * ```
 */
export interface DiagramNodeTextRun<S = unknown> {
	/** Run text content. */
	text: string;
	/**
	 * Raw parsed `a:rPr` run-properties object, preserved verbatim for
	 * round-trip. Untyped XML, hence the loose record shape.
	 */
	rPr?: Record<string, unknown>;
	/** Resolved standard shape-text style derived from {@link rPr}. */
	style?: S;
	/** Raw run XML used to retain unmodelled extension children on save. */
	rawXml?: Record<string, unknown>;
	/** Original direct-child order, including unmodelled extension children. */
	childOrder?: string[];
}

/** An ordered item within a SmartArt text paragraph. */
export type DiagramNodeTextParagraphItem<S = unknown> =
	| { kind: 'run'; run: DiagramNodeTextRun<S> }
	| {
			kind: 'break';
			rPr?: Record<string, unknown>;
			style?: S;
			rawXml?: Record<string, unknown>;
			childOrder?: string[];
	  }
	| {
			kind: 'field';
			id?: string;
			fieldType?: string;
			text: string;
			rPr?: Record<string, unknown>;
			style?: S;
			pPr?: Record<string, unknown>;
			rawXml?: Record<string, unknown>;
			childOrder?: string[];
	  }
	| { kind: 'tab'; rawXml?: Record<string, unknown>; childOrder?: string[] }
	| { kind: 'raw'; name: string; value: unknown };

/** A complete `a:p` paragraph in a SmartArt data-model text body. */
export interface DiagramNodeTextParagraph<S = unknown> {
	/** Paragraph properties (`a:pPr`) preserved verbatim. */
	pPr?: Record<string, unknown>;
	/** Text children in source order. */
	items: DiagramNodeTextParagraphItem<S>[];
	/** End-paragraph run properties (`a:endParaRPr`) preserved verbatim. */
	endParaRPr?: Record<string, unknown>;
	/** Resolved style for the paragraph terminator. */
	endParaStyle?: S;
	/** Raw paragraph XML used to retain unmodelled extension children on save. */
	rawXml?: Record<string, unknown>;
}

/**
 * Per-node visual override for a SmartArt node.
 *
 * Captures the individual fill / line / font colour and the bold / italic
 * emphasis a user has set on one specific node, independent of the diagram's
 * colour scheme and quick style. All colours are hex strings (e.g. "#FF0000").
 * Every field is optional: only the overridden aspects are carried, so an
 * empty object means "no per-node override".
 *
 * The parser reads these from the data point's `spPr` solid fill / line colour
 * and the first run's `rPr` (b / i / solidFill) when present, and the save path
 * writes them back so the override survives a load -> edit -> save round-trip.
 *
 * @example
 * ```ts
 * const style: DiagramNodeStyle = {
 *   fillColor: "#FF0000",
 *   fontColor: "#FFFFFF",
 *   bold: true,
 * };
 * // => satisfies DiagramNodeStyle
 * ```
 */
export interface DiagramNodeStyle {
	/** Solid fill colour override (hex, e.g. "#4F81BD"). */
	fillColor?: string;
	/** Outline / line colour override (hex). */
	lineColor?: string;
	/** Text (font) colour override (hex). */
	fontColor?: string;
	/** Bold emphasis override for the node's runs. */
	bold?: boolean;
	/** Italic emphasis override for the node's runs. */
	italic?: boolean;
}

/**
 * A single node in the SmartArt data model.
 *
 * @example
 * ```ts
 * const node: DiagramNode = {
 *   id: "1",
 *   text: "CEO",
 *   children: [
 *     { id: "2", text: "VP Marketing", parentId: "1" },
 *     { id: "3", text: "VP Engineering", parentId: "1" },
 *   ],
 * };
 * // => satisfies DiagramNode
 * ```
 */
export interface DiagramNode<S = unknown> {
	id: string;
	text: string;
	/** CT_Pt connection identifier, when the point references a connection. */
	connectionId?: string | null;
	parentId?: string;
	children?: DiagramNode<S>[];
	/** Node type from `@_type` attribute (e.g. "doc", "node", "asst", "pres"). */
	nodeType?: string;
	/**
	 * The node's own quick-style role (`dgm:prSet/@presStyleLbl` from its
	 * paired `type="pres"` presentation point, resolved via a `presOf`
	 * connection back to this content point). Structural names like `node1`,
	 * `asst2`, `bgShp`, `revTx`; distinct from {@link nodeType}, which is the
	 * data-model `@_type` ("node"/"asst"/...). Used to pick this node's own
	 * colour list from a colour transform's per-role palettes instead of the
	 * generic cycled palette (see `applySmartArtRoleColors`).
	 */
	styleRole?: string;
	/**
	 * `dgm:prSet/@coherent3DOff` (`CT_ElemPropSet`) resolved from the node's
	 * paired presentation point: when true, this node opts out of the
	 * diagram's overall coherent-3D scene rotation (a `dgm:scene3d`/`dgm:sp3d`
	 * quick-style variation) applied to every other node.
	 */
	coherent3DOff?: boolean;
	/**
	 * The layout variables PowerPoint recorded for each presentation node
	 * this point drives (`dgm:prSet/@presName` -> `dgm:presLayoutVars`
	 * variable -> `@val`), e.g. `{ hierRoot1: { hierBranch: 'l' } }` for an
	 * org-chart manager set to "Left Hanging". Read-only: the presentation
	 * points themselves round-trip verbatim.
	 */
	presLayoutVarsByName?: Record<string, Record<string, string>>;
	/**
	 * Per-run text + run-properties for the node's first paragraph, captured at
	 * parse time. When the joined run text still equals {@link text} (the node
	 * was not edited, or was edited only in ways that preserve the run split),
	 * the save path rebuilds the paragraph from these runs so per-run rich text
	 * is not flattened. When {@link text} diverges, the runs are ignored.
	 */
	runs?: DiagramNodeTextRun<S>[];
	/**
	 * Complete typed paragraph model. Unlike {@link runs}, this retains every
	 * paragraph and the ordered run, field, break, and tab children within it.
	 */
	paragraphs?: DiagramNodeTextParagraph<S>[];
	/**
	 * Optional per-node visual override (fill / line / font colour, bold /
	 * italic). Read at parse time from the point's `spPr` / first-run `rPr`, set
	 * by the editing op, honoured by the render path, and written back on save so
	 * it round-trips.
	 */
	style?: DiagramNodeStyle;
	/**
	 * Manual layout override read from the node's `dgm:prSet` `cust*`
	 * attributes (drag/resize/rotate/flip performed in PowerPoint's own diagram
	 * editor). Applied as a final transform after algorithmic layout by
	 * {@link module:smartart-layout-interpreter-custom} so it survives even
	 * when there is no cached `dsp:` drawing to fall back on.
	 */
	customLayout?: DiagramNodeCustomLayout;
}
