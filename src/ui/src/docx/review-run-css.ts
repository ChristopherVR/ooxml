import { wordHighlightColors } from 'ooxml-core/docx/ui';
import type { RunFormatting } from 'ooxml-core/docx';

/** Formatting normally supplied by direct marks, when Original mode neutralizes those marks. */
export function reviewRunCss(format: RunFormatting): string {
	const css: string[] = [];
	if (format.highlight && format.highlight !== 'none')
		css.push(`background-color:${wordHighlightColors[format.highlight]}`);
	else if (format.shadingFill && /^#[0-9a-f]{6}$/i.test(format.shadingFill))
		css.push(`background-color:${format.shadingFill}`);
	if (format.doubleStrike || format.underlineStyle === 'double')
		css.push('text-decoration-style:double');
	else if (format.underlineStyle && /dot/i.test(format.underlineStyle))
		css.push('text-decoration-style:dotted');
	else if (format.underlineStyle && /dash/i.test(format.underlineStyle))
		css.push('text-decoration-style:dashed');
	else if (format.underlineStyle && /wav/i.test(format.underlineStyle))
		css.push('text-decoration-style:wavy');
	if (format.underlineColor && /^#[0-9a-f]{6}$/i.test(format.underlineColor))
		css.push(`text-decoration-color:${format.underlineColor}`);
	return css.join(';');
}
