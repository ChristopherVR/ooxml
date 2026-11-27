/**
 * A deck's theme: the twelve colour slots, the font pair and a format scheme. The part itself is
 * written by the core's `drawingThemeXml`; only the format scheme fragment is authored here.
 */
import { drawingThemeXml } from '../../../../src/core/drawingml/write-theme';
import type { DrawingTheme, ThemeColorSlot } from '../../../../src/core/drawingml/theme-model';
import { FONTS } from '../../lib/fiction';

export interface DeckTheme {
	name: string;
	/** Hex colours (`#RRGGBB`) for the twelve theme slots. */
	colors: Record<ThemeColorSlot, string>;
	/** Whether the master maps the background to the dark slots (a dark deck). */
	dark: boolean;
}

/** Subtle fills, hairline to medium outlines and one soft shadow: nothing of the Office look. */
const FORMAT_SCHEME = (name: string) =>
	`<a:fmtScheme name="${name}">` +
	'<a:fillStyleLst>' +
	'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
	'<a:gradFill rotWithShape="1"><a:gsLst>' +
	'<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="108000"/><a:satMod val="104000"/></a:schemeClr></a:gs>' +
	'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="94000"/></a:schemeClr></a:gs>' +
	'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
	'<a:gradFill rotWithShape="1"><a:gsLst>' +
	'<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="116000"/><a:satMod val="108000"/></a:schemeClr></a:gs>' +
	'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="82000"/></a:schemeClr></a:gs>' +
	'</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>' +
	'</a:fillStyleLst>' +
	'<a:lnStyleLst>' +
	[6350, 12700, 25400]
		.map(
			(width) =>
				`<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`,
		)
		.join('') +
	'</a:lnStyleLst>' +
	'<a:effectStyleLst>' +
	'<a:effectStyle><a:effectLst/></a:effectStyle>' +
	'<a:effectStyle><a:effectLst/></a:effectStyle>' +
	'<a:effectStyle><a:effectLst><a:outerShdw blurRad="190500" dist="63500" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="18000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
	'</a:effectStyleLst>' +
	'<a:bgFillStyleLst>' +
	'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
	'<a:solidFill><a:schemeClr val="phClr"><a:tint val="96000"/></a:schemeClr></a:solidFill>' +
	'<a:gradFill rotWithShape="1"><a:gsLst>' +
	'<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="112000"/></a:schemeClr></a:gs>' +
	'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="78000"/></a:schemeClr></a:gs>' +
	'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
	'</a:bgFillStyleLst>' +
	'</a:fmtScheme>';

export function themeXml(theme: DeckTheme): string {
	const colors: DrawingTheme['colorScheme']['colors'] = {};
	for (const [slot, color] of Object.entries(theme.colors) as [ThemeColorSlot, string][])
		colors[slot] = { kind: 'srgb', value: color.replace('#', '').toUpperCase(), transforms: [] };
	const model: DrawingTheme = {
		name: theme.name,
		colorScheme: { name: theme.name, colors },
		fontScheme: {
			name: theme.name,
			major: { latin: FONTS.display, scripts: {} },
			minor: { latin: FONTS.body, scripts: {} },
		},
	};
	return drawingThemeXml(model, { formatScheme: FORMAT_SCHEME(theme.name) });
}
