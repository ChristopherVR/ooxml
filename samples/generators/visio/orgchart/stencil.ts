/** Styles and masters of the institute's drawing: position cards, and the timeline's bars. */
import { FONTS } from '../../lib/fiction';
import { connectorMaster } from '../lib/connector';
import { polygon, rectangle } from '../lib/geometry';
import { type MasterDef, MASTER_ROOT_ID } from '../lib/masters';
import { type Hue, FOG, HUES, INK, MIST, PAPER, SLATE, STONE } from '../lib/palette';
import { type PropDef, BOX_POINTS, PROP } from '../lib/shape';
import type { StyleDef } from '../lib/styles';
import { paint } from '../lib/theme';
import { colorCell } from '../lib/xml';

export const POSITION = { w: 1.75, h: 0.7 } as const;
/** The connection point a lead's team hangs from: just inside the card's lower left corner. */
export const TEAM_POINT = BOX_POINTS.length;
/** Points near the left end of a bar, where a dependency arrives: on top, then underneath. */
export const BAR_TOP = BOX_POINTS.length;

/** Shape Data of a position; the staff directory fills the same rows. */
export const POSITION_DATA: readonly PropDef[] = [
	{ name: 'Name', label: 'Name', type: PROP.string },
	{ name: 'Title', label: 'Title', type: PROP.string },
	{ name: 'Department', label: 'Department', type: PROP.string },
	{ name: 'Location', label: 'Location', type: PROP.string },
	{ name: 'Email', label: 'Email', type: PROP.string },
	{
		name: 'Started',
		label: 'Started',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
	{ name: 'FTE', label: 'Full-time equivalent', type: PROP.number, format: '0.0', value: 1 },
	{
		name: 'Status',
		label: 'Status',
		type: PROP.list,
		format: 'In post;Vacant;Seconded',
		value: 'In post',
	},
];

export const TASK_DATA: readonly PropDef[] = [
	{ name: 'Workstream', label: 'Workstream', type: PROP.string },
	{ name: 'Owner', label: 'Owner', type: PROP.string },
	{
		name: 'Start',
		label: 'Start',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
	{
		name: 'Finish',
		label: 'Finish',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
	{ name: 'Budget', label: 'Budget (thousands)', type: PROP.number, format: '0' },
	{
		name: 'Status',
		label: 'Status',
		type: PROP.list,
		format: 'Planned;Under way;Complete',
		value: 'Planned',
	},
];

export const MILESTONE_DATA: readonly PropDef[] = [
	{
		name: 'Date',
		label: 'Date',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
	{ name: 'Owner', label: 'Owner', type: PROP.string },
	{
		name: 'Status',
		label: 'Status',
		type: PROP.list,
		format: 'Planned;Reached;At risk',
		value: 'Planned',
	},
];

export const ORG_STYLES: readonly StyleDef[] = [
	{
		name: 'Position',
		line: { color: MIST, weight: 0.75, rounding: 0.05 },
		fill: { color: PAPER },
		text: { font: FONTS.display, size: 8.5, color: INK, align: 0, margin: 3 },
	},
	{ name: 'Band', line: { pattern: 0 }, fill: { color: HUES.sea.base } },
	{
		name: 'Reporting line',
		based: 'None',
		line: { color: STONE, weight: 1, pattern: 1 },
	},
	{
		name: 'Interval',
		line: { color: HUES.sea.base, weight: 0.75, rounding: 0.06 },
		fill: { color: HUES.sea.tint },
		text: { font: FONTS.body, size: 7.5, color: INK, align: 0, margin: 4 },
	},
	{
		name: 'Milestone',
		line: { color: PAPER, weight: 1 },
		fill: { color: HUES.amber.base },
		// The backdrop keeps a label readable where it crosses a grid line.
		text: { font: FONTS.display, size: 8, color: INK, margin: 0, background: PAPER },
	},
	{
		name: 'Dependency',
		based: 'None',
		line: { color: SLATE, weight: 0.75, pattern: 1, endArrow: 13, arrowSize: 1 },
	},
	{
		name: 'Row label',
		based: 'Text Only',
		text: { font: FONTS.display, size: 8.5, color: INK, align: 0, valign: 1, margin: 0 },
	},
	{
		name: 'Axis year',
		line: { pattern: 0 },
		fill: { color: HUES.sea.deep },
		text: { font: FONTS.display, size: 9, color: PAPER, margin: 2 },
	},
	{
		name: 'Axis quarter',
		line: { color: PAPER, weight: 0.75 },
		fill: { color: FOG },
		text: { font: FONTS.body, size: 7.5, color: SLATE, margin: 2 },
	},
	{ name: 'Today', line: { pattern: 0 }, fill: { color: HUES.coral.base } },
	{
		name: 'Today label',
		based: 'Text Only',
		text: { font: FONTS.display, size: 7.5, color: HUES.coral.deep, align: 0, valign: 1 },
	},
];

/** Local cells that give a card the colour of its department: the band's fill. */
export const department = (hue: Hue): Record<number, string[]> => ({
	[MASTER_ROOT_ID + 1]: [paint('FillForegnd', hue.base)],
});

/** Local cells that give a bar the colour of its workstream. */
export const workstream = (hue: Hue): string[] => [
	paint('LineColor', hue.base),
	colorCell('FillForegnd', hue.tint),
];

export const ORG_MASTERS: readonly MasterDef[] = [
	{
		name: 'Position',
		prompt: 'A post in the institute, filled or vacant.',
		size: { ...POSITION },
		shape: {
			style: 'Position',
			paths: [rectangle()],
			connections: [...BOX_POINTS, [[0, 0.1], 0]],
			props: POSITION_DATA,
			// Right of the band on the card's left edge.
			textBlock: { x: [0.5, 0.045], y: 0.5, w: [1, -0.13], h: 1 },
			children: [{ style: 'Band', paths: [rectangle()], box: { x: 0, y: 0, w: [0, 0.07], h: 1 } }],
		},
	},
	connectorMaster('Reporting line', 'Who a position reports to.', 'Reporting line'),
	{
		name: 'Interval',
		prompt: 'A piece of work with a start and a finish.',
		size: { w: 2, h: 0.36 },
		shape: {
			style: 'Interval',
			paths: [rectangle()],
			connections: [...BOX_POINTS, [[0, 0.15], 1], [[0, 0.15], 0]],
			props: TASK_DATA,
		},
	},
	{
		name: 'Milestone',
		prompt: 'A dated event with no duration.',
		size: { w: 0.26, h: 0.26 },
		shape: {
			style: 'Milestone',
			paths: [
				polygon([
					[0, 0.5],
					[0.5, 0],
					[1, 0.5],
					[0.5, 1],
				]),
			],
			connections: BOX_POINTS.slice(0, 4),
			props: MILESTONE_DATA,
			// The label hangs under the diamond.
			textBlock: { x: 0.5, y: [0, -0.24], w: [0, 1.8], h: [0, 0.36] },
		},
	},
	connectorMaster('Dependency', 'Work that must finish before other work starts.', 'Dependency'),
];
