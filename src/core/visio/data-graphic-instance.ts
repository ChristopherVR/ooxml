import {
	cell,
	esc,
	NATIVE_CALLOUTS,
	userRow,
	type NativeDataGraphicKind,
} from './data-graphic-masters';
import {
	VISIO_DATA_GRAPHIC_GAP as GAP,
	VISIO_DATA_GRAPHIC_STACK_STEP,
	visioDataGraphicCenter,
	visioDataGraphicPoint,
	type VisioDataGraphicPosition,
} from './data-graphic-place';

/** What one callout instance needs to be written. */
export interface CalloutInstance {
	id: string;
	/** The instance's name: the master's, then `Name.ID` for later ones. */
	name: string;
	masterId: string;
	ownerId: string;
	kind: NativeDataGraphicKind;
	/** The owner's Shape Data row the callout reads, and that row's label. */
	field: string;
	label: string;
	/** The row's value as saved: the number it holds, or its text. */
	value: { text: string; number?: number };
	position: VisioDataGraphicPosition;
	/** How many callouts already sit at this position, so this one stacks beyond them. */
	slot: number;
	owner: { width: number; height: number };
	min: number;
	max: number;
}

/** The callout's size by kind: text as wide as its owner, bars between 0.8 and 2 inches. */
export function calloutSize(
	kind: NativeDataGraphicKind,
	ownerWidth: number,
): { width: number; height: number } {
	if (kind === 'text') return { width: Math.max(ownerWidth, 1), height: 0.3 };
	if (kind === 'bar') return { width: Math.min(Math.max(ownerWidth, 0.8), 2), height: 0.16 };
	return { width: 0.22, height: 0.22 };
}

const ICON_COLORS = { low: '#e81123', mid: '#ffb900', high: '#107c10' };

/** The share of the range a value fills, clamped to 0..1. */
export const shareOf = (value: number, min: number, max: number): number =>
	Math.max(0, Math.min(1, (value - min) / (max - min || 1)));

const num = (value: number) => Number(value.toPrecision(10));

const rectRows = (width: number, height: number) =>
	`<Row T='LineTo' IX='2'>${cell('X', num(width), 'Inh')}</Row>` +
	`<Row T='LineTo' IX='3'>${cell('X', num(width), 'Inh')}${cell('Y', num(height), 'Inh')}</Row>` +
	`<Row T='LineTo' IX='4'>${cell('Y', num(height), 'Inh')}</Row>`;

/**
 * One callout as an instance of its master: it names the master, reads its value from the
 * owner's Shape Data row by formula (`Sheet.N!Prop.Row`), is placed beside the owner by a
 * formula on `User.visDGDefaultPos`, and caches what Visio would compute.
 */
export function calloutInstanceMarkup(input: CalloutInstance): string {
	const { kind, owner } = input;
	const size = calloutSize(kind, owner.width);
	const center = visioDataGraphicCenter(
		{ x: owner.width / 2, y: owner.height / 2, width: owner.width, height: owner.height },
		size,
		input.position,
		input.slot,
	);
	const stack = input.slot * VISIO_DATA_GRAPHIC_STACK_STEP;
	const point = visioDataGraphicPoint(input.position);
	const pos = `PNT(${point.x},${point.y})`;
	const reference = `Sheet.${input.ownerId}!Width`;
	const heightReference = `Sheet.${input.ownerId}!Height`;
	const pinX = `IF(PNTX(User.visDGDefaultPos)<0,-Width*0.5-${GAP},IF(PNTX(User.visDGDefaultPos)>4,${reference}+Width*0.5+${GAP},${reference}*0.5))`;
	const pinY = `IF(PNTY(User.visDGDefaultPos)>4,${heightReference}+Height*0.5+${GAP}+User.visDGStackHeight,IF(PNTY(User.visDGDefaultPos)<0,-Height*0.5-${GAP}-User.visDGStackHeight,${heightReference}*0.5-User.visDGStackHeight))`;
	const share =
		input.value.number === undefined ? 0 : shareOf(input.value.number, input.min, input.max);
	const numeric = input.value.number !== undefined;
	const field =
		`<Row N='msvCalloutField'>${cell('Value', input.value.text, `Sheet.${input.ownerId}!Prop.${input.field}`, numeric ? undefined : 'STR')}` +
		`${cell('Label', '')}${cell('SortKey', input.label)}</Row>`;
	const range =
		kind === 'text'
			? ''
			: `<Row N='msvCalloutPropMin'>${cell('Value', num(input.min))}</Row><Row N='msvCalloutPropMax'>${cell('Value', num(input.max))}</Row>`;
	const user =
		userRow('msvCalloutType', NATIVE_CALLOUTS[kind].type, undefined, 'STR') +
		userRow('visDGDefaultPos', pos, pos, 'PNT') +
		userRow('visDGStackHeight', num(stack)) +
		userRow('visDGField0', `{${input.field}}`, undefined, 'STR') +
		(kind === 'text' ? '' : `<Row N='PercentFilled'>${cell('Value', num(share), 'Inh')}</Row>`);
	let geometry = '';
	let text = '';
	if (kind === 'text') {
		geometry = `<Section N='Geometry' IX='0'>${rectRows(size.width, size.height)}</Section>`;
		text = `<Section N='Field'><Row IX='0'>${cell('Value', `${input.label}: ${input.value.text}`, 'Inh', 'STR')}</Row></Section>`;
	} else if (kind === 'bar') {
		geometry =
			`<Section N='Geometry' IX='0'>${rectRows(size.width, size.height)}</Section>` +
			`<Section N='Geometry' IX='1'>${rectRows(size.width * share, size.height)}</Section>`;
	} else {
		geometry = `<Section N='Geometry' IX='0'><Row T='Ellipse' IX='1'>${cell('X', num(size.width / 2), 'Inh')}${cell('Y', num(size.height / 2), 'Inh')}${cell('A', num(size.width), 'Inh')}${cell('B', num(size.height / 2), 'Inh')}${cell('C', num(size.width / 2), 'Inh')}${cell('D', num(size.height), 'Inh')}</Row></Section>`;
	}
	const color =
		share >= 2 / 3 ? ICON_COLORS.high : share >= 1 / 3 ? ICON_COLORS.mid : ICON_COLORS.low;
	return (
		`<Shape ID='${input.id}' NameU='${esc(input.name)}' Name='${esc(input.name)}' Type='Shape' Master='${input.masterId}'>` +
		cell('PinX', num(center.x), pinX) +
		cell('PinY', num(center.y), pinY) +
		cell('Width', num(size.width)) +
		cell('Height', num(size.height)) +
		cell('LocPinX', num(size.width / 2), 'Inh') +
		cell('LocPinY', num(size.height / 2), 'Inh') +
		(kind === 'icon' ? cell('FillForegnd', color, 'Inh') : '') +
		`<Section N='User'>${user}</Section>` +
		`<Section N='Property'>${field}${range}</Section>` +
		text +
		geometry +
		`</Shape>`
	);
}

/** The callout master's name for a kind, and the owner's User cells that make it a data graphic. */
export const calloutMasterName = (kind: NativeDataGraphicKind): string =>
	NATIVE_CALLOUTS[kind].name;
