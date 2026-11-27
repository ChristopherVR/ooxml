/** The flowchart's styles and masters: one master per kind of step, two kinds of connector. */
import { FONTS } from '../../lib/fiction';
import { connectorMaster } from '../lib/connector';
import { polygon, rectangle } from '../lib/geometry';
import type { MasterDef } from '../lib/masters';
import { type Hue, HUES, INK, PAPER, SLATE } from '../lib/palette';
import { type PropDef, BOX_POINTS, PROP } from '../lib/shape';
import type { StyleDef } from '../lib/styles';
import { cell } from '../lib/xml';

/** Shape Data every step carries; the linked recordset fills the same rows. */
export const STEP_DATA: readonly PropDef[] = [
	{ name: 'Step', label: 'Step', type: PROP.string, prompt: 'Reference in the handbook' },
	{ name: 'Owner', label: 'Owner', type: PROP.string, prompt: 'Role accountable for the step' },
	{ name: 'System', label: 'System', type: PROP.string, prompt: 'Where the work is recorded' },
	{ name: 'Duration', label: 'Duration (hours)', type: PROP.number, format: '0.0' },
	{ name: 'Cost', label: 'Cost per order', type: PROP.number, format: '0.00' },
	{
		name: 'Status',
		label: 'Status',
		type: PROP.list,
		format: 'In place;Under review;Planned',
		value: 'In place',
	},
	{
		name: 'Reviewed',
		label: 'Last reviewed',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
];

const kind = (name: string, hue: Hue): StyleDef => ({
	name,
	line: { color: hue.base, weight: 1 },
	fill: { color: hue.tint },
	text: { font: FONTS.body, size: 8.5, color: INK, margin: 3 },
});

export const FLOW_STYLES: readonly StyleDef[] = [
	{ ...kind('Step', HUES.sea), line: { color: HUES.sea.base, weight: 1, rounding: 0.06 } },
	kind('Decision', HUES.amber),
	kind('Document', HUES.plum),
	kind('Data', HUES.moss),
	kind('Reference', HUES.coral),
	{
		name: 'Terminal',
		line: { color: HUES.teal.base, weight: 1 },
		fill: { color: HUES.teal.base },
		text: { font: FONTS.display, size: 8.5, color: PAPER, margin: 3 },
	},
	{
		name: 'Sequence',
		based: 'None',
		line: { color: SLATE, weight: 1, pattern: 1, endArrow: 13, arrowSize: 2 },
		text: { font: FONTS.display, size: 7.5, color: INK, background: PAPER, margin: 1 },
	},
	{
		name: 'Message',
		based: 'Sequence',
		line: { color: HUES.plum.base, pattern: 2, endArrow: 13 },
	},
	{
		name: 'Phase label',
		based: 'Text Only',
		text: { font: FONTS.display, size: 8, color: SLATE, align: 0, valign: 1 },
	},
];

const W = 1.4;
const data = { connections: BOX_POINTS, props: STEP_DATA };
/** A text block inset from the outline, for shapes whose sides slope. */
const inset = (w: number, h: number) => ({ x: 0.5, y: 0.5, w, h });

export const FLOW_MASTERS: readonly MasterDef[] = [
	{
		name: 'Step',
		prompt: 'A piece of work done by one role.',
		size: { w: W, h: 0.7 },
		shape: { style: 'Step', paths: [rectangle()], ...data },
	},
	{
		name: 'Decision',
		prompt: 'A question with two outcomes.',
		size: { w: W, h: 0.9 },
		shape: {
			style: 'Decision',
			paths: [
				polygon([
					[0, 0.5],
					[0.5, 0],
					[1, 0.5],
					[0.5, 1],
				]),
			],
			textBlock: inset(0.66, 0.6),
			...data,
			connections: BOX_POINTS.slice(0, 4),
		},
	},
	{
		name: 'Start or end',
		prompt: 'Where the process begins or stops.',
		size: { w: W, h: 0.56 },
		shape: {
			style: 'Terminal',
			paths: [rectangle()],
			// A stadium: the corner radius follows half the height.
			cells: [cell('Rounding', 0.28, 'Height*0.5')],
			...data,
		},
	},
	{
		name: 'Document',
		prompt: 'A document the step produces or reads.',
		size: { w: W, h: 0.72 },
		shape: {
			style: 'Document',
			paths: [
				{
					segments: [
						{ t: 'M', x: 0, y: 1 },
						{ t: 'L', x: 1, y: 1 },
						{ t: 'L', x: 1, y: 0.16 },
						// The torn lower edge: one S curve from right to left.
						{ t: 'C', x: 0, y: 0.16, x1: 0.68, y1: 0.46, x2: 0.32, y2: -0.16 },
						{ t: 'L', x: 0, y: 1 },
					],
				},
			],
			textBlock: { x: 0.5, y: 0.6, w: 1, h: 0.8 },
			...data,
			connections: [[0.5, 1], [1, 0.5], [0.5, 0.15], [0, 0.5], ...BOX_POINTS.slice(4)],
		},
	},
	{
		name: 'Data',
		prompt: 'Data read from or written to a system.',
		size: { w: W, h: 0.7 },
		shape: {
			style: 'Data',
			paths: [
				polygon([
					[0.14, 0],
					[1, 0],
					[0.86, 1],
					[0, 1],
				]),
			],
			textBlock: inset(0.72, 1),
			...data,
			connections: [[0.5, 1], [0.93, 0.5], [0.5, 0], [0.07, 0.5], ...BOX_POINTS.slice(4)],
		},
	},
	{
		name: 'Off-page reference',
		prompt: 'Continues on another page.',
		size: { w: 1.0, h: 0.72 },
		shape: {
			style: 'Reference',
			paths: [
				polygon([
					[0, 1],
					[1, 1],
					[1, 0.36],
					[0.5, 0],
					[0, 0.36],
				]),
			],
			textBlock: { x: 0.5, y: 0.64, w: 1, h: 0.72 },
			connections: BOX_POINTS.slice(0, 4),
		},
	},
	connectorMaster('Sequence flow', 'The order in which work is done.', 'Sequence'),
	connectorMaster('Message flow', 'A document or notice sent to another role.', 'Message'),
];
