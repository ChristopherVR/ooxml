/**
 * The document theme: a DrawingML theme part with the Visio theme extensions ([MS-VSDX] 2.2.7.4),
 * holding the samples' palette and fonts, and the PageSheet cells that select it. The core
 * library writes only its built-in themes and under extension URIs of its own, so this part is
 * written here with the URIs Visio reads.
 */
import { FONTS } from '../../lib/fiction';
import { FOG, HUES, INK, PAPER } from './palette';
import { XML_HEADER, cell, colorCell } from './xml';

/** The scheme number the pages select; Visio's own themes use other numbers. */
export const THEME_SCHEME = 60;
const VT = 'http://schemas.microsoft.com/office/visio/2012/theme';

/** Theme colour names as THEMEVAL() takes them, with the colour each has in this theme. */
export const THEME_COLORS = {
	AccentColor: HUES.teal.base,
	AccentColor2: HUES.sea.base,
	AccentColor3: HUES.amber.base,
	AccentColor4: HUES.coral.base,
	AccentColor5: HUES.moss.base,
	AccentColor6: HUES.plum.base,
} as const;
export type ThemeColor = keyof typeof THEME_COLORS;

/** The formula of a cell that follows a theme colour. */
export const themed = (name: ThemeColor): string => `THEMEGUARD(THEMEVAL("${name}"))`;

export const isThemeColor = (color: string): boolean =>
	Object.values(THEME_COLORS).some((value) => value.toLowerCase() === color.toLowerCase());

/**
 * A colour cell. One of the theme's accent colours is saved as its theme formula with the colour
 * as the cached value, so the shape follows the theme in Visio; any other colour is literal.
 */
export function paint(name: string, color: string): string {
	const entry = Object.entries(THEME_COLORS).find(
		([, value]) => value.toLowerCase() === color.toLowerCase(),
	);
	return entry
		? cell(name, color.toLowerCase(), themed(entry[0] as ThemeColor))
		: colorCell(name, color);
}

const hex = (color: string) => color.slice(1).toUpperCase();
const srgb = (color: string) => `<a:srgbClr val="${hex(color)}"/>`;
const schemeId = (namespace = '') =>
	`<vt:schemeID${namespace} schemeEnum="${THEME_SCHEME}" schemeGUID="{00000000-0000-0000-0000-000000000000}"/>`;
const scheme = schemeId();
const rootScheme = schemeId(` xmlns:vt="${VT}"`);
const ext = (uri: string, body: string) => `<a:ext uri="${uri}">${body}</a:ext>`;
const repeat = (count: number, text: string) => Array.from({ length: count }, () => text).join('');
const PLACEHOLDER = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
const line = (points: number) =>
	`<a:ln w="${Math.round(points * 12700)}" cap="flat" cmpd="sng">${PLACEHOLDER}<a:prstDash val="solid"/></a:ln>`;
const NO_EFFECT = '<a:effectStyle><a:effectLst/></a:effectStyle>';

/** Fill, line, effect and background lists: flat fills, plain lines, no effects. */
function formats(styles: number, widths: readonly number[]): string {
	return (
		`<a:fillStyleLst>${repeat(styles, PLACEHOLDER)}</a:fillStyleLst>` +
		`<a:lnStyleLst>${widths.map(line).join('')}</a:lnStyleLst>` +
		`<a:effectStyleLst>${repeat(styles, NO_EFFECT)}</a:effectStyleLst>` +
		`<a:bgFillStyleLst>${repeat(3, PLACEHOLDER)}</a:bgFillStyleLst>`
	);
}

const font = (name: string) =>
	`<a:latin typeface="${name}"/><a:ea typeface=""/><a:cs typeface=""/>`;
const fontProps = (color: string) =>
	`<vt:fontProps style="0"><vt:color><a:schemeClr val="${color}"/></vt:color></vt:fontProps>`;
const lineStyle = (end: number) =>
	`<vt:lineStyle><vt:lineEx rndg="0" start="0" startSize="2" end="${end}" endSize="2" pattern="1"/></vt:lineStyle>`;

/**
 * The four colour variants of the theme (Design > Variants): seven colours each. The first is
 * the one the pages select, in the order THEMEVAL("AccentColor"), ("AccentColor2")... reads them.
 */
const VARIANTS: readonly (readonly string[])[] = (() => {
	const { teal, sea, amber, coral, moss, plum } = HUES;
	const base = [teal, sea, amber, coral, moss, plum].map((hue) => hue.base);
	const rotate = (by: number) => [...base.slice(by), ...base.slice(0, by)];
	return [0, 1, 3, 5].map((by) => [...rotate(by), INK]);
})();
const variation = (colors: readonly string[]) =>
	`<vt:variationClrScheme>${colors.map((color, index) => `<vt:varColor${index + 1}>${srgb(color)}</vt:varColor${index + 1}>`).join('')}</vt:variationClrScheme>`;
const varStyle = (index: number) =>
	`<vt:varStyle fillIdx="${index}" lineIdx="${index}" effectIdx="${index}" fontIdx="${index}"/>`;

/** `visio/theme/theme1.xml`. */
export function themeXml(name: string): string {
	const colors =
		`<a:dk1>${srgb(INK)}</a:dk1><a:lt1>${srgb(PAPER)}</a:lt1>` +
		`<a:dk2>${srgb(HUES.sea.deep)}</a:dk2><a:lt2>${srgb(FOG)}</a:lt2>` +
		Object.values(THEME_COLORS)
			.map((color, index) => `<a:accent${index + 1}>${srgb(color)}</a:accent${index + 1}>`)
			.join('') +
		`<a:hlink>${srgb(HUES.sea.base)}</a:hlink><a:folHlink>${srgb(HUES.plum.base)}</a:folHlink>`;
	const colorScheme =
		`<a:clrScheme name="${name}">${colors}<a:extLst>` +
		ext('{2703A3B3-D2E1-43D9-8057-6E9D74E0F44A}', rootScheme) +
		ext(
			'{093E89EA-6996-430E-BFF9-83A9FAAAAB73}',
			`<vt:bkgnd xmlns:vt="${VT}">${srgb(PAPER)}</vt:bkgnd>`,
		) +
		ext(
			'{DDD2D869-C2EF-471E-B8FA-914AFA308C9F}',
			`<vt:variationClrSchemeLst xmlns:vt="${VT}">${VARIANTS.map(variation).join('')}</vt:variationClrSchemeLst>`,
		) +
		'</a:extLst></a:clrScheme>';
	const fontScheme =
		`<a:fontScheme name="${name}"><a:majorFont>${font(FONTS.display)}</a:majorFont>` +
		`<a:minorFont>${font(FONTS.body)}</a:minorFont><a:extLst>` +
		ext('{B3C0788F-89CC-4CDB-A57D-0664CFDA821B}', rootScheme) +
		'</a:extLst></a:fontScheme>';
	const formatScheme = `<a:fmtScheme name="${name}">${formats(6, [0.75, 1, 1.5, 0.75, 1, 1.5])}</a:fmtScheme>`;
	const extensions =
		ext(
			'{1342405F-259F-4C95-8CDF-A9DCE2D418A2}',
			`<vt:fmtConnectorScheme xmlns:vt="${VT}" name="${name}">${formats(3, [0.75, 1, 1.5])}</vt:fmtConnectorScheme>`,
		) +
		ext(
			'{D75FF423-9257-4291-A4FE-1B2448832E17}',
			`<vt:themeScheme xmlns:vt="${VT}">${scheme}</vt:themeScheme>`,
		) +
		ext(
			'{27CD58D4-7086-4B73-B2AB-7AEBE2148A8C}',
			`<vt:fmtSchemeEx xmlns:vt="${VT}">${scheme}</vt:fmtSchemeEx>`,
		) +
		ext(
			'{C6430689-8E98-42EC-ADBF-087148533A3F}',
			`<vt:fmtConnectorSchemeEx xmlns:vt="${VT}">${scheme}</vt:fmtConnectorSchemeEx>`,
		) +
		ext(
			'{56243398-1771-4C39-BF73-A5702A9C147F}',
			`<vt:fillStyles xmlns:vt="${VT}">${repeat(6, '<vt:fillProps pattern="1"/>')}</vt:fillStyles>`,
		) +
		ext(
			'{6CAB99AB-0A78-4BAB-B597-526D62367CB4}',
			`<vt:lineStyles xmlns:vt="${VT}"><vt:fmtConnectorSchemeLineStyles>${lineStyle(0)}${repeat(3, lineStyle(13))}</vt:fmtConnectorSchemeLineStyles>` +
				`<vt:fmtSchemeLineStyles>${repeat(7, lineStyle(0))}</vt:fmtSchemeLineStyles></vt:lineStyles>`,
		) +
		ext(
			'{EBE24D50-EC5C-4D6F-A1A3-C5F0A18B936A}',
			`<vt:fontStylesGroup xmlns:vt="${VT}"><vt:connectorFontStyles>${repeat(3, fontProps('dk1'))}</vt:connectorFontStyles>` +
				`<vt:fontStyles>${repeat(3, fontProps('dk1'))}${repeat(3, fontProps('lt1'))}</vt:fontStyles></vt:fontStylesGroup>`,
		) +
		ext(
			'{494CE47F-D151-47DC-95E8-85652EA8A67E}',
			`<vt:variationStyleSchemeLst xmlns:vt="${VT}">${repeat(4, `<vt:variationStyleScheme embellishment="1">${[1, 2, 4, 6].map(varStyle).join('')}</vt:variationStyleScheme>`)}</vt:variationStyleSchemeLst>`,
		);
	return (
		`${XML_HEADER}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}">` +
		`<a:themeElements>${colorScheme}${fontScheme}${formatScheme}<a:extLst>${extensions}</a:extLst></a:themeElements>` +
		'<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
	);
}

/** The PageSheet cells that make a page use the theme. */
export const THEME_PAGE_CELLS = [
	'ColorSchemeIndex',
	'EffectSchemeIndex',
	'ConnectorSchemeIndex',
	'FontSchemeIndex',
	'ThemeIndex',
]
	.map((name) => cell(name, THEME_SCHEME))
	.concat(cell('VariationColorIndex', 0), cell('VariationStyleIndex', 0))
	.join('');

export const sampleTheme = (name: string): { xml: string; pageCells: string } => ({
	xml: themeXml(name),
	pageCells: THEME_PAGE_CELLS,
});
