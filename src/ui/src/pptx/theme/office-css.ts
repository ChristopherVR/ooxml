import { LEGACY_OFFICE_THEME_TOKENS, legacyThemeAliases, OFFICE_TOKENS } from '../../theme';
const MARKER = '/* ooxml-pptx-office-theme */';
/** Resolve legacy inputs at the actual chrome scope, without declaring inherited legacy defaults. */
export function withOfficeThemeCss(css: string, selector: string): string {
	if (css.includes(MARKER)) return css;
	function reference(value: string, start: number) {
		let depth = 1,
			end = start + 4;
		for (; end < value.length && depth; end++) {
			if (value[end] === '(') depth++;
			else if (value[end] === ')') depth--;
		}
		if (depth) return undefined;
		const body = value.slice(start + 4, end - 1);
		const comma = body.indexOf(',');
		return {
			name: (comma < 0 ? body : body.slice(0, comma)).trim(),
			fallback: comma < 0 ? '' : body.slice(comma + 1).trim(),
			end,
		};
	}
	function rewrite(value: string): string {
		let result = '',
			cursor = 0;
		while (cursor < value.length) {
			const start = value.indexOf('var(', cursor);
			if (start < 0) return result + value.slice(cursor);
			const current = reference(value, start);
			if (!current) return result + value.slice(cursor);
			result += value.slice(cursor, start);
			const key = current.name.slice('--pptx-'.length);
			if (current.name.startsWith('--pptx-') && key in LEGACY_OFFICE_THEME_TOKENS) {
				const names = [current.name];
				let fallback = current.fallback;
				while (fallback.startsWith('var(')) {
					const next = reference(fallback, 0);
					if (
						!next ||
						next.end !== fallback.length ||
						!next.name.startsWith('--pptx-') ||
						!(next.name.slice(7) in LEGACY_OFFICE_THEME_TOKENS)
					)
						break;
					names.push(next.name);
					fallback = next.fallback;
				}
				const canonical = `var(--_pptx-${key}${fallback ? ', ' + rewrite(fallback) : ''})`;
				result +=
					names.length > 1
						? names.reduceRight((inner, name) => `var(${name}, ${inner})`, canonical)
						: canonical;
			} else
				result += `var(${current.name}${current.fallback ? ', ' + rewrite(current.fallback) : ''})`;
			cursor = current.end;
		}
		return result;
	}
	return `${MARKER}\n${legacyThemeAliases(selector, '--pptx-', '--_pptx-')}\n${rewrite(css)}`;
}
/** A legacy override followed by the shared Office token at the point of use. */
export function pptxThemeValue(name: keyof typeof LEGACY_OFFICE_THEME_TOKENS): string {
	const token = LEGACY_OFFICE_THEME_TOKENS[name];
	return `var(--pptx-${name}, var(${token}, ${OFFICE_TOKENS[token]}))`;
}
