/** Preserve XML attribute whitespace which some DOM serializers emit literally.
 * Comments, CDATA, processing instructions and text are deliberately left untouched.
 */
export function preserveAttributeWhitespace(xml: string): string {
	return xml.replace(
		/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?[^!?\s>][^"'>]*(?:(?:"[^"]*"|'[^']*')[^"'>]*)*>/g,
		(token) => {
			if (token.startsWith('<!') || token.startsWith('<?')) return token;
			return token.replace(/"[^"]*"|'[^']*'/g, (value) =>
				value.replace(/[\t\r\n]/g, (character) => `&#${character.charCodeAt(0)};`),
			);
		},
	);
}
