/**
 * Run formatting in Word's clipboard HTML that CSS cannot say, read from its `mso-` and Office
 * HTML properties (`mso-highlight`, `text-underline`, `text-line-through`, `mso-hide`,
 * `mso-text-raise`) and from the CSS a browser would read differently from Word.
 */
import type { TextRun } from '../model-run';
import { wordHighlightColors, type WordHighlightToken } from '../highlight';
import { isWordUnderlineStyle } from '../underline';
import { halfPoints, signedTwips } from '../units';
import { cssPoints } from './css';

/** The run properties the editor carries as one `data-run-props` mark. */
export type PastedRunProperties = Pick<
	TextRun,
	| 'caps'
	| 'smallCaps'
	| 'doubleStrike'
	| 'vanish'
	| 'underlineStyle'
	| 'characterSpacingTwips'
	| 'positionHalfPoints'
>;

/** `mso-highlight` names (the sixteen HTML colour names) and Word's highlight tokens. */
const HIGHLIGHTS: Record<string, Exclude<WordHighlightToken, 'none'>> = {
	black: 'black',
	blue: 'blue',
	aqua: 'cyan',
	cyan: 'cyan',
	lime: 'green',
	fuchsia: 'magenta',
	magenta: 'magenta',
	red: 'red',
	yellow: 'yellow',
	white: 'white',
	navy: 'darkBlue',
	teal: 'darkCyan',
	green: 'darkGreen',
	purple: 'darkMagenta',
	maroon: 'darkRed',
	olive: 'darkYellow',
	gray: 'darkGray',
	silver: 'lightGray',
};

/** The HTML 4 colour names Word writes for `color`, which the editor reads only as hex. */
const NAMED_COLORS: Record<string, string> = {
	black: '#000000',
	white: '#ffffff',
	red: '#ff0000',
	lime: '#00ff00',
	blue: '#0000ff',
	yellow: '#ffff00',
	aqua: '#00ffff',
	fuchsia: '#ff00ff',
	maroon: '#800000',
	green: '#008000',
	navy: '#000080',
	olive: '#808000',
	purple: '#800080',
	teal: '#008080',
	gray: '#808080',
	silver: '#c0c0c0',
};

export interface WordRunStyle {
	/** The CSS left for the browser-facing marks (fonts, colour, highlight). */
	css: Map<string, string>;
	props: PastedRunProperties;
	/** A `<s>` that Word means as a double strikethrough is not a single strike. */
	dropStrike: boolean;
}

function underlineStyle(value: string | undefined): PastedRunProperties['underlineStyle'] {
	if (!value) return undefined;
	const token = value
		.trim()
		.split(/\s+/)[0]!
		.replace(/-([a-z])/g, (_m, char: string) => char.toUpperCase());
	return token !== 'single' && token !== 'none' && isWordUnderlineStyle(token) ? token : undefined;
}

/** Reads one element's inline style; `mso-` declarations never reach the output CSS. */
export function wordRunStyle(declarations: ReadonlyMap<string, string>): WordRunStyle {
	const css = new Map<string, string>();
	const props: PastedRunProperties = {};
	for (const [name, value] of declarations) {
		if (name.startsWith('mso-') || name === 'background' || name === 'background-color') continue;
		if (name === 'text-underline' || name === 'text-line-through' || name === 'display') continue;
		if (name === 'color') {
			const lower = value.toLowerCase();
			if (lower === 'windowtext' || lower === 'auto') continue;
			css.set('color', NAMED_COLORS[lower] ?? value);
			continue;
		}
		css.set(name, value);
	}
	const highlight = HIGHLIGHTS[declarations.get('mso-highlight')?.toLowerCase() ?? ''];
	if (highlight) css.set('background-color', wordHighlightColors[highlight]);
	const underline = underlineStyle(declarations.get('text-underline'));
	if (underline) props.underlineStyle = underline;
	const doubleStrike = declarations.get('text-line-through')?.toLowerCase() === 'double';
	if (doubleStrike) props.doubleStrike = true;
	const variant = declarations.get('font-variant')?.toLowerCase();
	if (variant === 'small-caps') props.smallCaps = true;
	if (declarations.get('text-transform')?.toLowerCase() === 'uppercase') props.caps = true;
	if (
		declarations.get('display')?.toLowerCase() === 'none' ||
		declarations.get('mso-hide')?.toLowerCase() === 'all'
	)
		props.vanish = true;
	const spacing = cssPoints(declarations.get('letter-spacing'));
	if (spacing) props.characterSpacingTwips = signedTwips(Math.round(spacing * 20));
	const raise = cssPoints(declarations.get('mso-text-raise'));
	if (raise) props.positionHalfPoints = halfPoints(Math.round(raise * 2));
	css.delete('font-variant');
	css.delete('text-transform');
	css.delete('letter-spacing');
	return { css, props, dropStrike: doubleStrike };
}
