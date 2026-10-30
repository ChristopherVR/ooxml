// Character formatting: the .doc run properties as `w:rPr` children, written only where the run
// differs from what its paragraph and character styles already give it.
import type { DocCharProps, DocUnderlineKind } from '@christophervr/ole2/doc-char-props';
import type { Conversion } from './context';
import { el } from './xml';

const UNDERLINE: Record<DocUnderlineKind, string> = {
	none: 'none',
	single: 'single',
	words: 'words',
	double: 'double',
	dotted: 'dotted',
	thick: 'thick',
	dash: 'dash',
	dotDash: 'dotDash',
	dotDotDash: 'dotDotDash',
	wave: 'wave',
	dottedHeavy: 'dottedHeavy',
	dashHeavy: 'dashedHeavy',
	dotDashHeavy: 'dashDotHeavy',
	dotDotDashHeavy: 'dashDotDotHeavy',
	waveHeavy: 'wavyHeavy',
	dashLong: 'dashLong',
	dashLongHeavy: 'dashLongHeavy',
	waveDouble: 'wavyDouble',
};

type Toggle = 'bold' | 'italic' | 'caps' | 'smallCaps' | 'strike' | 'doubleStrike' | 'outline';
const TOGGLES: [Toggle | 'boldCs' | 'italicCs', string][] = [
	['bold', 'w:b'],
	['boldCs', 'w:bCs'],
	['italic', 'w:i'],
	['italicCs', 'w:iCs'],
	['caps', 'w:caps'],
	['smallCaps', 'w:smallCaps'],
	['strike', 'w:strike'],
	['doubleStrike', 'w:dstrike'],
	['outline', 'w:outline'],
];
const TOGGLES_AFTER: [keyof DocCharProps, string][] = [
	['shadow', 'w:shadow'],
	['emboss', 'w:emboss'],
	['imprint', 'w:imprint'],
	['noProof', 'w:noProof'],
	['hidden', 'w:vanish'],
];

const toggle = (
	props: DocCharProps,
	base: DocCharProps,
	key: keyof DocCharProps,
	tag: string,
): string => {
	const value = props[key] as boolean | undefined;
	if (value === undefined || value === Boolean(base[key])) return '';
	return value ? el(tag) : el(tag, { 'w:val': '0' });
};

/** The `w:rPr` children of a run (without the wrapper), in schema order. */
export function runPropertiesXml(
	props: DocCharProps,
	base: DocCharProps,
	conversion: Conversion,
	options: { style?: boolean } = {},
): string {
	let out = '';
	const style = options.style === false ? undefined : props.styleId;
	if (style !== undefined && !conversion.styles.isDefaultCharacter(style)) {
		const id = conversion.styles.idOf(style);
		if (id) out += el('w:rStyle', { 'w:val': id });
	}
	const names = props.fonts ?? {};
	const baseNames = base.fonts ?? {};
	const fonts = {
		'w:ascii': names.ascii !== baseNames.ascii ? names.ascii : undefined,
		'w:hAnsi': names.hAnsi !== baseNames.hAnsi ? names.hAnsi : undefined,
		'w:eastAsia': names.eastAsia !== baseNames.eastAsia ? names.eastAsia : undefined,
		'w:cs': names.cs !== baseNames.cs ? names.cs : undefined,
	};
	if (Object.values(fonts).some((v) => v !== undefined)) out += el('w:rFonts', fonts);
	for (const [key, tag] of TOGGLES) out += toggle(props, base, key, tag);
	for (const [key, tag] of TOGGLES_AFTER) out += toggle(props, base, key, tag);
	if (props.color && props.color !== base.color) out += el('w:color', { 'w:val': props.color });
	if (props.spacingTwips !== undefined && props.spacingTwips !== base.spacingTwips)
		out += el('w:spacing', { 'w:val': props.spacingTwips });
	if (props.scalePercent !== undefined && props.scalePercent !== base.scalePercent)
		out += el('w:w', { 'w:val': props.scalePercent });
	if (props.kerningPt !== undefined && props.kerningPt !== base.kerningPt)
		out += el('w:kern', { 'w:val': Math.round(props.kerningPt * 2) });
	if (props.positionPt !== undefined && props.positionPt !== base.positionPt)
		out += el('w:position', { 'w:val': Math.round(props.positionPt * 2) });
	if (props.sizePt !== undefined && props.sizePt !== base.sizePt)
		out += el('w:sz', { 'w:val': Math.round(props.sizePt * 2) });
	if (props.sizeCsPt !== undefined && props.sizeCsPt !== base.sizeCsPt)
		out += el('w:szCs', { 'w:val': Math.round(props.sizeCsPt * 2) });
	if (props.highlight && props.highlight !== base.highlight)
		out += el('w:highlight', { 'w:val': props.highlight });
	if (props.underline && props.underline !== base.underline)
		out += el('w:u', {
			'w:val': UNDERLINE[props.underline],
			'w:color': props.underlineColor,
		});
	if (props.verticalAlign && props.verticalAlign !== base.verticalAlign)
		out += el('w:vertAlign', { 'w:val': props.verticalAlign });
	if (props.rtl !== undefined && props.rtl !== Boolean(base.rtl))
		out += props.rtl ? el('w:rtl') : el('w:rtl', { 'w:val': '0' });
	return out;
}

export const wrapRunProperties = (inner: string): string =>
	inner ? `<w:rPr>${inner}</w:rPr>` : '';
