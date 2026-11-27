/**
 * The architecture drawing's styles and masters. Every node is a card (a group master): a white
 * tile, a pictogram on its left and two lines of text. A node takes the colour of its tier by
 * overriding the pictogram's line and fill.
 */
import { FONTS } from '../../lib/fiction';
import { connectorMaster } from '../lib/connector';
import { type Path, rectangle } from '../lib/geometry';
import { type MasterDef, MASTER_ROOT_ID } from '../lib/masters';
import { type Hue, HUES, INK, MIST, PAPER, SLATE } from '../lib/palette';
import { type PropDef, BOX_POINTS, PROP } from '../lib/shape';
import type { StyleDef } from '../lib/styles';
import { paint } from '../lib/theme';
import { colorCell } from '../lib/xml';
import * as glyphs from './glyphs';

export const CARD = { w: 2.1, h: 0.8 } as const;
/** The pictogram is the card's only sub-shape. */
const GLYPH_ID = MASTER_ROOT_ID + 1;

/** Shape Data of every node; the inventory recordset fills the same rows. */
export const NODE_DATA: readonly PropDef[] = [
	{ name: 'Hostname', label: 'Host name', type: PROP.string },
	{ name: 'IP', label: 'IP address', type: PROP.string },
	{ name: 'Role', label: 'Role', type: PROP.string },
	{ name: 'Owner', label: 'Owning team', type: PROP.string },
	{
		name: 'Status',
		label: 'Status',
		type: PROP.list,
		format: 'Live;Degraded;Planned',
		value: 'Live',
	},
	{ name: 'Cost', label: 'Monthly cost', type: PROP.number, format: '0.00' },
	{
		name: 'Commissioned',
		label: 'Commissioned',
		type: PROP.date,
		format: '{{d MMM yyyy}}',
		value: '2026-01-15T00:00:00',
	},
];

const flow = (name: string, color: string, extra: StyleDef['line'] = {}): StyleDef => ({
	name,
	based: 'None',
	line: { color, weight: 1.25, pattern: 1, endArrow: 13, arrowSize: 2, ...extra },
	text: { font: FONTS.display, size: 7.5, color: SLATE, background: PAPER, margin: 1 },
});

export const ARCHITECTURE_STYLES: readonly StyleDef[] = [
	{
		name: 'Card',
		line: { color: MIST, weight: 0.75, rounding: 0.06 },
		fill: { color: PAPER },
		text: { font: FONTS.display, size: 8.5, color: INK, align: 0, margin: 2 },
	},
	{
		name: 'Glyph',
		line: { color: HUES.sea.base, weight: 1 },
		fill: { color: HUES.sea.tint },
	},
	flow('Data flow', HUES.teal.base),
	flow('Control flow', HUES.amber.base, { pattern: 2 }),
	flow('Network link', SLATE, { endArrow: 0, weight: 1.5 }),
	flow('Replication', HUES.plum.base, { pattern: 2 }),
];

/** The local cells that give a node's pictogram the colour of its tier. */
export const tier = (hue: Hue): Record<number, string[]> => ({
	[GLYPH_ID]: [paint('LineColor', hue.base), colorCell('FillForegnd', hue.tint)],
});

function card(name: string, prompt: string, glyph: readonly Path[]): MasterDef {
	return {
		name,
		prompt,
		size: { ...CARD },
		shape: {
			style: 'Card',
			paths: [rectangle()],
			connections: BOX_POINTS,
			props: NODE_DATA,
			// The text sits right of the pictogram.
			textBlock: { x: [0.5, 0.39], y: 0.5, w: [1, -0.92], h: 1 },
			children: [
				{
					style: 'Glyph',
					paths: glyph,
					box: { x: [0, 0.16], y: [0.5, -0.26], w: [0, 0.52], h: [0, 0.52] },
				},
			],
		},
	};
}

export const ARCHITECTURE_MASTERS: readonly MasterDef[] = [
	card('Weather station', 'An Aerie station in the field.', glyphs.STATION),
	card('Home hub', 'The hub that relays the stations of one home.', glyphs.HUB),
	card('Gateway', 'Terminates connections at the edge of a tier.', glyphs.GATEWAY),
	card('Queue', 'Holds messages until a consumer takes them.', glyphs.QUEUE),
	card('Service', 'A deployable service.', glyphs.SERVICE),
	card('Database', 'A database.', glyphs.DATABASE),
	card('Archive', 'Object storage for files kept long term.', glyphs.ARCHIVE),
	card('Web client', 'A browser application.', glyphs.WEB),
	card('Mobile client', 'A phone application.', glyphs.MOBILE),
	card('Partner system', 'A system run by another organisation.', glyphs.PARTNER),
	card('Server', 'A virtual machine.', glyphs.SERVER),
	card('Firewall', 'Filters traffic between networks.', glyphs.FIREWALL),
	card('Load balancer', 'Spreads requests over several nodes.', glyphs.BALANCER),
	card('Network cloud', 'A network outside the platform.', glyphs.CLOUD),
	connectorMaster('Data flow', 'Readings and results moving through the platform.', 'Data flow'),
	connectorMaster('Control flow', 'Commands and notifications.', 'Control flow'),
	connectorMaster('Network link', 'A network connection between two hosts.', 'Network link'),
	connectorMaster('Replication', 'Data copied to a standby.', 'Replication'),
];
