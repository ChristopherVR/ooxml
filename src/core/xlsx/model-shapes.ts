// Worksheet shapes and form controls. Both are kept verbatim for saving (`sourceXml`); the
// types here are what the painter and the form-control interactions read.
import type {
	DrawingColor,
	DrawingFill,
	DrawingFrame,
	DrawingLine,
	DrawingPath,
	DrawingTextBody,
} from '../drawingml/types';
import type { DrawingAnchor } from './model';

/** A DrawingML shape, connector or group of a worksheet drawing, with its `xdr:style` references. */
export interface ShapeNode {
	kind: 'shape' | 'connector' | 'group';
	name?: string;
	/** `txBox="1"`: a text box (no style-derived fill or line). */
	textBox?: boolean;
	/** Position and size in EMU, in the coordinates of the parent (the group's child space). */
	frame: DrawingFrame;
	/** Degrees clockwise. */
	rotation?: number;
	flipH?: boolean;
	flipV?: boolean;
	/** Preset name (`rect`, `rightArrow`...) or `custom`. */
	geometry: string;
	adjustments?: Record<string, number>;
	paths?: DrawingPath[];
	fill?: DrawingFill;
	line?: DrawingLine;
	/** `xdr:style` colours: `fillRef`, `lnRef` and `fontRef`. */
	style?: {
		fill?: { idx: number; color?: DrawingColor };
		line?: { idx: number; color?: DrawingColor };
		font?: { color?: DrawingColor };
	};
	text?: DrawingTextBody;
	/** A group's children and the child coordinate space they are drawn in. */
	children?: ShapeNode[];
	childFrame?: DrawingFrame;
}

export type FormControlType =
	| 'checkbox'
	| 'radio'
	| 'dropdown'
	| 'list'
	| 'button'
	| 'spin'
	| 'scroll'
	| 'groupBox'
	| 'label'
	| 'edit'
	| 'other';

/**
 * A legacy form control: its drawing anchor (kept verbatim), the `formControlPr` part and the
 * caption of its VML text box. State changes made in the UI go through the linked cell; the
 * `formControlPr` attributes (`checked`, `sel`, `val`) are rewritten on save.
 */
export interface FormControlObject {
	kind: 'formControl';
	anchor: DrawingAnchor;
	name?: string;
	/** `cNvPr/@id`, which `<control shapeId>` and the VML `_x0000_s<id>` shape share. */
	shapeId: number;
	controlType: FormControlType;
	/** Part name of the `formControlPr`, when the sheet references one. */
	partName?: string;
	caption?: string;
	/** `fmlaLink`, as written (`$C$1`, `Sheet1!$C$1`). */
	linkedCell?: string;
	/** `fmlaRange`: the list of a drop-down or list box. */
	listRange?: string;
	checked?: 'checked' | 'mixed' | 'unchecked';
	/** First option button of a group (`firstButton`). */
	firstButton?: boolean;
	/** 1-based selected list item (`sel`). */
	selected?: number;
	value?: number;
	min?: number;
	max?: number;
	step?: number;
	page?: number;
	horizontal?: boolean;
	/** `dropStyle`: `combo`, `comboedit` or `simple`. */
	dropStyle?: string;
	/** The anchor element as read (the `mc:AlternateContent` holding the control shape). */
	sourceXml: string;
	/** Linked cell or state changed since loading: the save rewrites the `formControlPr` part. */
	edited?: true;
}
