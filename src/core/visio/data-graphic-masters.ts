import { NS, parseXml } from '../xml/index';
import { applyShapeCreationStyles } from './edit-shape-create';
import type { MasterSpec } from './edit-master-add';

/** The callout kinds a native data graphic is made of. */
export type NativeDataGraphicKind = 'text' | 'bar' | 'icon';

export interface NativeCalloutDefinition {
	/** The master's name: Visio's own name for this kind of callout. */
	name: string;
	/** What Visio's `User.msvCalloutType` says for this kind. */
	type: 'Text Callout' | 'Data Bar' | 'Icon Set';
	uniqueId: string;
}

/**
 * The three callout masters this package authors. They carry the names and the `msvCallout*`
 * cells Visio's own callouts use (`Prop.msvCalloutField` is the data the callout shows), but the
 * shapes are this package's own, much simpler than Microsoft's DVCALL stencil.
 */
export const NATIVE_CALLOUTS: Record<NativeDataGraphicKind, NativeCalloutDefinition> = {
	text: {
		name: 'Text callout',
		type: 'Text Callout',
		uniqueId: '{4F0C7A21-0001-0000-8E40-00608CF305B2}',
	},
	bar: {
		name: 'Data bar 1',
		type: 'Data Bar',
		uniqueId: '{4F0C7A21-0002-0000-8E40-00608CF305B2}',
	},
	icon: {
		name: 'Status icons',
		type: 'Icon Set',
		uniqueId: '{4F0C7A21-0003-0000-8E40-00608CF305B2}',
	},
};

export const esc = (text: string): string =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&apos;');

/** `<Cell>` markup: value, optional unit and formula. */
export function cell(
	name: string,
	value: string | number,
	formula?: string,
	unit?: string,
): string {
	return `<Cell N='${name}' V='${esc(String(value))}'${unit ? ` U='${unit}'` : ''}${
		formula === undefined ? '' : ` F='${esc(formula)}'`
	}/>`;
}

/** A User row: `[name, value, formula?, unit?]`. */
export const userRow = (name: string, value: string | number, formula?: string, unit?: string) =>
	`<Row N='${name}'>${cell('Value', value, formula, unit)}${cell('Prompt', '', 'No Formula')}</Row>`;

const REAL_PROPERTY = (name: string, value: string, label: string, sort: string, type: number) =>
	`<Row N='${name}'>${cell('Value', value)}${cell('Prompt', '')}${cell('Label', label)}${cell('Format', '')}${cell('SortKey', sort)}${cell('Type', type)}${cell('Invisible', 1)}${cell('Verify', 0)}${cell('DataLinked', 0, 'No Formula')}${cell('LangID', 'en-US')}${cell('Calendar', 0)}</Row>`;

const RECT = (width: string, height: string, section = 0, noFill = 0, noLine = 0) =>
	`<Section N='Geometry' IX='${section}'>${cell('NoFill', noFill)}${cell('NoLine', noLine)}${cell('NoShow', 0)}${cell('NoSnap', 0)}${cell('NoQuickDrag', 0, 'No Formula')}` +
	`<Row T='MoveTo' IX='1'>${cell('X', 0, 'Width*0')}${cell('Y', 0, 'Height*0')}</Row>` +
	`<Row T='LineTo' IX='2'>${cell('X', 0, width)}${cell('Y', 0, 'Height*0')}</Row>` +
	`<Row T='LineTo' IX='3'>${cell('X', 0, width)}${cell('Y', 0, height)}</Row>` +
	`<Row T='LineTo' IX='4'>${cell('X', 0, 'Width*0')}${cell('Y', 0, height)}</Row>` +
	`<Row T='LineTo' IX='5'>${cell('X', 0, 'Width*0')}${cell('Y', 0, 'Height*0')}</Row></Section>`;

/** The cells every callout master shares: pin, size, no handles. */
const frame = (width: number, height: number) =>
	cell('PinX', width / 2) +
	cell('PinY', height / 2) +
	cell('Width', width) +
	cell('Height', height) +
	cell('LocPinX', width / 2, 'Width*0.5') +
	cell('LocPinY', height / 2, 'Height*0.5') +
	cell('Angle', 0) +
	cell('FlipX', 0) +
	cell('FlipY', 0) +
	cell('ResizeMode', 0) +
	cell('NoObjHandles', 1) +
	cell('LockCalcWH', 1);

const PERCENT =
	'MAX(0,MIN(1,(Prop.msvCalloutField-Prop.msvCalloutPropMin)/(Prop.msvCalloutPropMax-Prop.msvCalloutPropMin)))';

/** The master's shape, by kind (shape ID 5, as Visio numbers the first shape of a master). */
function shapeXml(kind: NativeDataGraphicKind): string {
	const { type } = NATIVE_CALLOUTS[kind];
	const user = (rows: string) =>
		`<Section N='User'>${userRow('msvCalloutType', type, undefined, 'STR')}${userRow('visDGCalloutItem', 5)}${userRow('visDGStackHeight', 0)}${userRow('visVersion', 15)}${rows}</Section>`;
	const head = `<Shape ID='5' Type='Shape'>`;
	if (kind === 'text')
		return (
			head +
			frame(1.25, 0.3) +
			cell('LinePattern', 1) +
			cell('LineColor', '#7f7f7f') +
			cell('FillPattern', 1) +
			cell('FillForegnd', '#ffffff') +
			user(userRow('visDGDefaultPos', 0)) +
			`<Section N='Property'>${REAL_PROPERTY('msvCalloutField', '0', '', 'Label', 0)}</Section>` +
			`<Section N='Character'><Row IX='0'>${cell('Size', 0.1111111111111111, undefined, 'PT')}</Row></Section>` +
			RECT('Width*1', 'Height*1') +
			`<Section N='Field'><Row IX='0'>${cell('Value', 'Label: 0', 'Prop.msvCalloutField.SortKey&": "&Prop.msvCalloutField', 'STR')}${cell('Format', 'esc(0)', 'FIELDPICTURE(0)', 'STR')}${cell('Type', 0)}${cell('UICat', 0, 'No Formula')}${cell('UICod', 0, 'No Formula')}${cell('UIFmt', 0, 'No Formula')}${cell('Calendar', 0)}${cell('ObjectKind', 0)}</Row></Section>` +
			`<Text><cp IX='0'/><pp IX='0'/><fld IX='0'>Label: 0</fld></Text></Shape>`
		);
	const properties =
		REAL_PROPERTY('msvCalloutPropMin', '0', 'Minimum Value', '', 2) +
		REAL_PROPERTY('msvCalloutPropMax', '100', 'Maximum Value', '', 2) +
		REAL_PROPERTY('msvCalloutField', '75', '', 'Label', 2);
	const percent = userRow('PercentFilled', 0.75, PERCENT);
	if (kind === 'bar')
		return (
			head +
			frame(1.25, 0.16) +
			cell('LinePattern', 1) +
			cell('LineColor', '#7f7f7f') +
			cell('FillPattern', 1) +
			cell('FillForegnd', '#2e75b6') +
			user(userRow('visDGDefaultPos', 0) + percent) +
			`<Section N='Property'>${properties}</Section>` +
			// The frame is an outline; the bar is the second path, as wide as the share of the range.
			RECT('Width*1', 'Height*1', 0, 1, 0) +
			RECT('Width*User.PercentFilled', 'Height*1', 1, 0, 1) +
			`</Shape>`
		);
	return (
		head +
		frame(0.22, 0.22) +
		cell('LinePattern', 0) +
		cell('FillPattern', 1) +
		cell(
			'FillForegnd',
			'#107c10',
			'IF(User.PercentFilled>=0.6667,RGB(16,124,16),IF(User.PercentFilled>=0.3333,RGB(255,185,0),RGB(232,17,35)))',
		) +
		user(userRow('visDGDefaultPos', 0) + userRow('PercentFilled', 0.75, PERCENT)) +
		`<Section N='Property'>${properties}</Section>` +
		`<Section N='Geometry' IX='0'>${cell('NoFill', 0)}${cell('NoLine', 0)}${cell('NoShow', 0)}${cell('NoSnap', 0)}${cell('NoQuickDrag', 0, 'No Formula')}` +
		`<Row T='Ellipse' IX='1'>${cell('X', 0.11, 'Width*0.5')}${cell('Y', 0.11, 'Height*0.5')}${cell('A', 0.22, 'Width*1')}${cell('B', 0.11, 'Height*0.5')}${cell('C', 0.11, 'Width*0.5')}${cell('D', 0.22, 'Height*1')}</Row></Section></Shape>`
	);
}

const SIZE: Record<NativeDataGraphicKind, { width: number; height: number }> = {
	text: { width: 1.25, height: 0.3 },
	bar: { width: 1.25, height: 0.16 },
	icon: { width: 0.22, height: 0.22 },
};

/** Parse markup under the drawing's namespace and give it to `document`. */
export function importMarkup(document: Document, namespace: string, markup: string): Element {
	const parsed = parseXml(`<Wrapper xmlns="${namespace}">${markup}</Wrapper>`).documentElement;
	const first = Array.from(parsed.childNodes).find((node) => node.nodeType === 1)!;
	return document.importNode(first, true) as Element;
}

const ENTRY_ATTRIBUTES = (
	id: string,
	name: string,
	uniqueId: string,
	hidden: boolean,
	type?: string,
) =>
	[
		['ID', id],
		['NameU', name],
		['IsCustomNameU', '1'],
		['Name', name],
		['IsCustomName', '1'],
		['Prompt', ''],
		['IconSize', hidden ? '0' : '1'],
		['AlignName', hidden ? '0' : '2'],
		['MatchByName', '0'],
		['IconUpdate', '0'],
		['UniqueID', uniqueId],
		['BaseID', uniqueId.replace(/^\{[0-9A-F]{8}/i, '{4F0C7A10')],
		['PatternFlags', '0'],
		...(hidden ? [['Hidden', '1']] : []),
		...(type ? [['MasterType', type]] : []),
	] as const;

function appendEntry(
	masters: Element,
	attributes: ReturnType<typeof ENTRY_ATTRIBUTES>,
	pageSheet: string,
	relationshipId: string,
): void {
	const doc = masters.ownerDocument!;
	const master = doc.createElementNS(masters.namespaceURI, 'Master');
	for (const [key, value] of attributes) master.setAttribute(key, value);
	master.appendChild(
		importMarkup(doc, masters.namespaceURI!, `<PageSheet>${pageSheet}</PageSheet>`),
	);
	master
		.appendChild(doc.createElementNS(masters.namespaceURI, 'Rel'))
		.setAttributeNS(NS.r, 'r:id', relationshipId);
	masters.appendChild(master);
}

const pageCells = (size: { width: number; height: number }) =>
	cell('PageWidth', size.width) +
	cell('PageHeight', size.height) +
	cell('PageScale', 1, undefined, 'IN_F') +
	cell('DrawingScale', 1, undefined, 'IN_F') +
	cell('DrawingSizeType', 0) +
	cell('DrawingScaleType', 0);

function contentsRoot(namespace: string, document: Document | Element, shapes: string): Element {
	void document;
	const root = parseXml(
		`<MasterContents xmlns="${namespace}"><Shapes>${shapes}</Shapes></MasterContents>`,
	).documentElement;
	root.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
	return root;
}

/** The spec that adds one callout master to a drawing. */
export function calloutMasterSpec(kind: NativeDataGraphicKind): MasterSpec {
	const definition = NATIVE_CALLOUTS[kind];
	return {
		uniqueId: definition.uniqueId,
		name: definition.name,
		entry: (masters, id, name, relationshipId) =>
			appendEntry(
				masters,
				ENTRY_ATTRIBUTES(id, name, definition.uniqueId, false),
				pageCells(SIZE[kind]),
				relationshipId,
			),
		contents: (document) => {
			const root = contentsRoot(document.namespaceURI!, document, shapeXml(kind));
			const shape = root.getElementsByTagNameNS(root.namespaceURI, 'Shape')[0]!;
			applyShapeCreationStyles(shape, document);
			return root;
		},
	};
}

/** What a data graphic master remembers about one item. */
export interface NativeMasterItem {
	kind: NativeDataGraphicKind;
	field: string;
	position: string;
}

/**
 * The spec of the hidden data graphic master Visio keeps for each graphic: its page sheet holds
 * the graphic's options (`visDGNextItemID`, `msvDGCalloutGap`), its shapes one entry per item
 * naming the callout master (`visDGUseMaster`) and the field it shows (`visDGField0`).
 */
export function dataGraphicMasterSpec(
	uniqueId: string,
	name: string,
	items: readonly NativeMasterItem[],
): MasterSpec {
	const entries = items
		.map((item, index) => {
			const callout = NATIVE_CALLOUTS[item.kind];
			return (
				`<Shape ID='${6 + index}' NameU='${esc(callout.name)}' Name='${esc(callout.name)}' Type='Shape'>` +
				`<Section N='User'>${userRow('msvCalloutType', callout.type, undefined, 'STR')}${userRow('visDGDefaultPos', item.position, item.position, 'PNT')}${userRow('visDGItemID', index + 1)}${userRow('visDGUseMaster', 254, `USE("${callout.name}")`)}${userRow('visDGField0', `{${item.field}}`, undefined, 'STR')}${userRow('visDGCalloutField', `{${item.field}}`, 'User.visDGField0', 'STR')}</Section></Shape>`
			);
		})
		.join('');
	return {
		uniqueId,
		name,
		entry: (masters, id, unique, relationshipId) =>
			appendEntry(
				masters,
				ENTRY_ATTRIBUTES(id, unique, uniqueId, true, '2'),
				pageCells({ width: 8.5, height: 11 }) +
					`<Section N='User'>${userRow('visUSEType', 1)}${userRow('visDGVisible', 1, 'TRUE', 'BOOL')}${userRow('visDGShowBorder', 0, undefined, 'BOOL')}${userRow('visDGHideText', 0, undefined, 'BOOL')}${userRow('visDGNextItemID', items.length + 2)}${userRow('msvDGCalloutGap', 0.05)}</Section>`,
				relationshipId,
			),
		contents: (document) =>
			contentsRoot(
				document.namespaceURI!,
				document,
				`<Shape ID='5' Type='Group'>${cell('PinX', 0)}${cell('PinY', 0)}${cell('Width', 1)}${cell('Height', 1)}${cell('LocPinX', 0, 'Width*0')}${cell('LocPinY', 0, 'Height*0')}<Shapes>${entries}</Shapes></Shape>`,
			),
	};
}
