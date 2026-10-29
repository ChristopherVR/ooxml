// Legacy form controls (check box, option button, drop-down, list box, button, spin button, scroll
// bar, group box, label). The drawing part holds each control as an `mc:AlternateContent` anchor
// around a hidden shape; the sheet's `<controls>` link the shape id to a `formControlPr` part, and
// the legacy VML drawing holds the caption. The XML of every part is kept verbatim for saving.
import { NS, elements, first, parseXml, relAttr, type XmlElement } from '../../xml/index';
import type { DrawingAnchor, DrawingObject, FormControlObject, Worksheet } from '../model';
import type { FormControlType } from '../model-shapes';
import type { SourceIndex } from './package';
import { att, selfContainedXml } from './xml-util';

const TYPES: Record<string, FormControlType> = {
	checkbox: 'checkbox',
	radio: 'radio',
	drop: 'dropdown',
	list: 'list',
	button: 'button',
	spin: 'spin',
	scroll: 'scroll',
	gbox: 'groupBox',
	label: 'label',
	edit: 'edit',
};

/** Reads a drawing-level `mc:AlternateContent` anchor that holds a form control shape. */
export function readFormControlFrame(
	alternate: XmlElement,
	anchorOf: (anchor: XmlElement) => DrawingAnchor,
): FormControlObject | undefined {
	const choice = elements(alternate).find((node) => node.localName === 'Choice');
	const anchorEl = choice
		? elements(choice).find((node) => node.localName?.endsWith('Anchor'))
		: undefined;
	if (!anchorEl) return undefined;
	const shape = elements(anchorEl).find((node) => node.localName === 'sp');
	const nv = first(shape, 'nvSpPr', NS.xdr);
	const cNvPr = first(nv, 'cNvPr', NS.xdr);
	const spid = [...(cNvPr?.getElementsByTagName('*') ?? [])].find(
		(node) => node.localName === 'compatExt',
	);
	if (!spid) return undefined;
	const shapeId = Number(att(cNvPr, 'id'));
	if (!Number.isFinite(shapeId)) return undefined;
	const name = att(cNvPr, 'name');
	return {
		kind: 'formControl',
		anchor: anchorOf(anchorEl),
		shapeId,
		controlType: 'other',
		...(name ? { name } : {}),
		sourceXml: selfContainedXml(alternate),
	};
}

const num = (value: string | undefined): number | undefined => {
	if (value === undefined || value === '') return undefined;
	const parsed = Number(value);
	return Number.isFinite(parsed) ? parsed : undefined;
};

function readProps(control: FormControlObject, xml: string): void {
	const props = parseXml(xml, { label: 'XLSX form control' }).documentElement;
	control.controlType = TYPES[(att(props, 'objectType') ?? '').toLowerCase()] ?? 'other';
	const link = att(props, 'fmlaLink');
	if (link) control.linkedCell = link;
	const range = att(props, 'fmlaRange');
	if (range) control.listRange = range;
	if (control.controlType === 'checkbox' || control.controlType === 'radio') {
		const checked = att(props, 'checked');
		control.checked =
			checked === 'Checked' ? 'checked' : checked === 'Mixed' ? 'mixed' : 'unchecked';
	}
	if (att(props, 'firstButton') === '1') control.firstButton = true;
	for (const [key, attribute] of [
		['selected', 'sel'],
		['value', 'val'],
		['min', 'min'],
		['max', 'max'],
		['step', 'inc'],
		['page', 'page'],
	] as const) {
		const value = num(att(props, attribute));
		if (value !== undefined) control[key] = value;
	}
	if (att(props, 'horiz') === '1') control.horizontal = true;
	const style = att(props, 'dropStyle');
	if (style) control.dropStyle = style;
}

const entities: Record<string, string> = {
	amp: '&',
	lt: '<',
	gt: '>',
	quot: '"',
	apos: "'",
	nbsp: ' ',
};

/** The caption of the VML shape `_x0000_s<id>`: the text of its `v:textbox`. */
export function vmlCaption(vml: string, shapeId: number): string | undefined {
	const start = vml.indexOf(`id="_x0000_s${shapeId}"`);
	if (start < 0) return undefined;
	const end = vml.indexOf('</v:shape>', start);
	const shape = vml.slice(start, end < 0 ? undefined : end);
	const box = /<v:textbox[^>]*>([\s\S]*?)<\/v:textbox>/.exec(shape)?.[1];
	if (!box) return undefined;
	const text = box
		.replace(/<br\s*\/?>/gi, '\n')
		.replace(/<[^>]*>/g, '')
		.replace(/&(#x?[0-9a-f]+|\w+);/gi, (match, code: string) => {
			if (code.startsWith('#')) {
				const value =
					code[1]?.toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : Number(code.slice(1));
				return Number.isFinite(value) ? String.fromCodePoint(value) : match;
			}
			return entities[code.toLowerCase()] ?? match;
		})
		.replace(/\s+/g, ' ')
		.trim();
	return text || undefined;
}

/**
 * Completes the form controls found in the drawing with the type and state of their
 * `formControlPr` part and their VML caption. A control the sheet does not list stays `other`.
 */
export function attachFormControls(
	source: SourceIndex,
	sheetPart: string,
	root: XmlElement,
	sheet: Worksheet,
	legacyDrawing: XmlElement | undefined,
): void {
	const controls = sheet.drawings.filter(
		(object): object is FormControlObject => object.kind === 'formControl',
	);
	if (!controls.length) return;
	const rels = source.rels(sheetPart);
	const parts = new Map<number, string>();
	for (const control of Array.from(root.getElementsByTagNameNS(NS.x, 'control'))) {
		const shapeId = Number(att(control, 'shapeId'));
		const rel = rels.get(relAttr(control, 'id') ?? '');
		const target = rel ? source.target(sheetPart, rel) : undefined;
		if (target) parts.set(shapeId, target);
	}
	const vmlRel = legacyDrawing ? rels.get(relAttr(legacyDrawing, 'id') ?? '') : undefined;
	const vmlPart = vmlRel ? source.target(sheetPart, vmlRel) : undefined;
	const vml = vmlPart ? source.text(vmlPart) : undefined;
	for (const control of controls) {
		const part = parts.get(control.shapeId);
		const xml = part ? source.text(part) : undefined;
		if (part && xml) {
			control.partName = part;
			try {
				readProps(control, xml);
			} catch {
				control.controlType = 'other';
			}
		}
		const caption = vml ? vmlCaption(vml, control.shapeId) : undefined;
		if (caption) control.caption = caption;
	}
}

export const isFormControl = (object: DrawingObject): object is FormControlObject =>
	object.kind === 'formControl';
