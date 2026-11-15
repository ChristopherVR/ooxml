/** WordprocessingML ST_HighlightColor tokens, including `none` for clearing. */
export const WORD_HIGHLIGHT_TOKENS = [
	'black',
	'blue',
	'cyan',
	'green',
	'magenta',
	'red',
	'yellow',
	'white',
	'darkBlue',
	'darkCyan',
	'darkGreen',
	'darkMagenta',
	'darkRed',
	'darkYellow',
	'darkGray',
	'lightGray',
	'none',
] as const;

export type WordHighlightToken = (typeof WORD_HIGHLIGHT_TOKENS)[number];

export function isWordHighlightToken(value: string): value is WordHighlightToken {
	return (WORD_HIGHLIGHT_TOKENS as readonly string[]).includes(value);
}
/** Word's highlight palette tokens mapped to stable browser colors. */
export const wordHighlightColors = {
	black: '#000000',
	blue: '#0000ff',
	cyan: '#00ffff',
	green: '#00ff00',
	magenta: '#ff00ff',
	red: '#ff0000',
	yellow: '#ffff00',
	white: '#ffffff',
	darkBlue: '#000080',
	darkCyan: '#008080',
	darkGreen: '#008000',
	darkMagenta: '#800080',
	darkRed: '#800000',
	darkYellow: '#808000',
	darkGray: '#808080',
	lightGray: '#c0c0c0',
} satisfies Record<Exclude<WordHighlightToken, 'none'>, string>;
