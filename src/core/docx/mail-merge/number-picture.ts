import { formatValue } from '../../xlsx/numfmt/format';

/** Adapts Word's numeric picture tokens to the shared spreadsheet number renderer. */
export function formatMergeNumber(value: number, picture: string): string {
	const sections: string[][] = [[]];
	for (const token of picture.match(/'[^']*'|[\s\S]/g) ?? []) {
		if (token === ';') sections.push([]);
		else sections[sections.length - 1]?.push(token);
	}
	const selected = value < 0 ? sections[1] : value === 0 ? sections[2] : undefined;
	const tokens = selected ?? sections[0] ?? [];
	const point = tokens.indexOf('.');
	const integer = tokens.slice(0, point === -1 ? undefined : point);
	const truncateAt = integer.findIndex((token) => token === 'x' || token === 'X');
	const truncateTo =
		truncateAt === -1
			? undefined
			: integer.slice(truncateAt).filter((token) => /^[xX0#]$/.test(token)).length;
	const group = integer.includes(',');
	const explicitSign = tokens.includes('-') || tokens.includes('+');
	const escape = (literal: string) => Array.from(literal, (character) => `\\${character}`).join('');
	const format = tokens
		.map((token, index) => {
			if (token.startsWith("'") && token.length > 1) return escape(token.slice(1, -1));
			if (truncateTo !== undefined && index < truncateAt && /^[0#]$/.test(token)) return '';
			if (token === '#') return group || truncateTo !== undefined ? '#' : '?';
			if (token === 'x' || token === 'X') return point === -1 || index < point ? '#' : '0';
			if (token === '0' || token === '.') return token;
			if (token === ',')
				return integer.slice(0, index).some((prior) => /^[0#xX]$/.test(prior)) ? ',' : '#,';
			if (token === '+' || token === '-')
				return escape(value < 0 ? '-' : value > 0 && token === '+' ? '+' : ' ');
			return escape(token);
		})
		.join('');
	if (!format) return '';
	const magnitude = Math.abs(value);
	const number = truncateTo === undefined ? magnitude : magnitude % 10 ** truncateTo;
	return formatValue(value < 0 && !selected && !explicitSign ? -number : number, format).text;
}
