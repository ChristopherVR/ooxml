// Paragraph formatting: the .doc paragraph properties as `w:pPr` children, written only where the
// paragraph differs from what its style already gives it.
import type { DocBorders, DocParaProps, DocTabStop } from '@christophervr/ole2/doc-para-props';
import { borderXml, bordersEqual, shadingXml } from './borders';
import { el } from './xml';

const JUSTIFY: Record<string, string> = {
	left: 'left',
	center: 'center',
	right: 'right',
	justify: 'both',
	distribute: 'distribute',
};
const TAB_ALIGN: Record<DocTabStop['alignment'], string> = {
	left: 'left',
	center: 'center',
	right: 'right',
	decimal: 'decimal',
	bar: 'bar',
};
const TAB_LEADER: Record<DocTabStop['leader'], string | undefined> = {
	none: undefined,
	dot: 'dot',
	hyphen: 'hyphen',
	underscore: 'underscore',
	heavy: 'heavy',
	middleDot: 'middleDot',
};

type Flag =
	| 'keepWithNext'
	| 'keepTogether'
	| 'pageBreakBefore'
	| 'widowControl'
	| 'noLineNumbers'
	| 'noAutoHyphen'
	| 'bidi'
	| 'contextualSpacing'
	| 'mirrorIndents';

const flag = (props: DocParaProps, base: DocParaProps, key: Flag, tag: string): string => {
	const value = props[key];
	if (value === undefined || value === Boolean(base[key])) return '';
	return value ? el(tag) : el(tag, { 'w:val': '0' });
};

function bordersXml(props: DocBorders | undefined, base: DocBorders | undefined): string {
	if (!props) return '';
	let out = '';
	for (const side of ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const)
		if (!bordersEqual(props[side], base?.[side])) out += borderXml(side, props[side]);
	return out ? `<w:pBdr>${out}</w:pBdr>` : '';
}

function tabsXml(props: DocParaProps, base: DocParaProps): string {
	const tabs = props.tabs ?? [];
	const inherited = base.tabs ?? [];
	if (JSON.stringify(tabs) === JSON.stringify(inherited)) return '';
	let out = '';
	for (const tab of inherited)
		if (!tabs.some((t) => t.positionTwips === tab.positionTwips))
			out += el('w:tab', { 'w:val': 'clear', 'w:pos': tab.positionTwips });
	for (const tab of tabs) {
		const same = inherited.find((t) => JSON.stringify(t) === JSON.stringify(tab));
		if (same) continue;
		out += el('w:tab', {
			'w:val': TAB_ALIGN[tab.alignment],
			'w:leader': TAB_LEADER[tab.leader],
			'w:pos': tab.positionTwips,
		});
	}
	return out ? `<w:tabs>${out}</w:tabs>` : '';
}

/** The `w:pPr` children of a paragraph or style (no wrapper, no style reference), in schema order. */
export function paragraphPropertiesXml(
	props: DocParaProps,
	base: DocParaProps,
	options: { numbering?: boolean } = {},
): string {
	let out = '';
	out += flag(props, base, 'keepWithNext', 'w:keepNext');
	out += flag(props, base, 'keepTogether', 'w:keepLines');
	out += flag(props, base, 'pageBreakBefore', 'w:pageBreakBefore');
	out += flag(props, base, 'widowControl', 'w:widowControl');
	if (options.numbering !== false) {
		const ilfo = props.ilfo ?? 0;
		const level = props.ilvl ?? 0;
		const same = ilfo === (base.ilfo ?? 0) && (ilfo === 0 || level === (base.ilvl ?? 0));
		if (!same && ilfo >= 0)
			out += `<w:numPr>${level ? el('w:ilvl', { 'w:val': level }) : ''}${el('w:numId', { 'w:val': ilfo })}</w:numPr>`;
	}
	out += flag(props, base, 'noLineNumbers', 'w:suppressLineNumbers');
	out += bordersXml(props.borders, base.borders);
	if (JSON.stringify(props.shading) !== JSON.stringify(base.shading))
		out += shadingXml(props.shading);
	out += tabsXml(props, base);
	out += flag(props, base, 'noAutoHyphen', 'w:suppressAutoHyphens');
	out += flag(props, base, 'bidi', 'w:bidi');
	const spacing = spacingXml(props, base);
	if (spacing) out += spacing;
	const indent = indentXml(props, base);
	if (indent) out += indent;
	out += flag(props, base, 'contextualSpacing', 'w:contextualSpacing');
	out += flag(props, base, 'mirrorIndents', 'w:mirrorIndents');
	if (props.alignment && props.alignment !== base.alignment)
		out += el('w:jc', { 'w:val': JUSTIFY[props.alignment] });
	if (props.outlineLevel !== base.outlineLevel)
		out += el('w:outlineLvl', { 'w:val': props.outlineLevel ?? 9 });
	return out;
}

function spacingXml(props: DocParaProps, base: DocParaProps): string {
	const before =
		props.spaceBeforeTwips !== base.spaceBeforeTwips ? props.spaceBeforeTwips : undefined;
	const after = props.spaceAfterTwips !== base.spaceAfterTwips ? props.spaceAfterTwips : undefined;
	const beforeAuto = props.spaceBeforeAuto && !base.spaceBeforeAuto ? '1' : undefined;
	const afterAuto = props.spaceAfterAuto && !base.spaceAfterAuto ? '1' : undefined;
	const line =
		JSON.stringify(props.lineSpacing) !== JSON.stringify(base.lineSpacing)
			? props.lineSpacing
			: undefined;
	if (before === undefined && after === undefined && !line && !beforeAuto && !afterAuto) return '';
	return el('w:spacing', {
		'w:before': before,
		'w:beforeAutospacing': beforeAuto,
		'w:after': after,
		'w:afterAutospacing': afterAuto,
		'w:line': line?.value,
		'w:lineRule': line?.rule,
	});
}

function indentXml(props: DocParaProps, base: DocParaProps): string {
	const left = props.indentLeftTwips !== base.indentLeftTwips ? props.indentLeftTwips : undefined;
	const right =
		props.indentRightTwips !== base.indentRightTwips ? props.indentRightTwips : undefined;
	const first = props.firstLineTwips !== base.firstLineTwips ? props.firstLineTwips : undefined;
	if (left === undefined && right === undefined && first === undefined) return '';
	return el('w:ind', {
		'w:left': left,
		'w:right': right,
		'w:firstLine': first !== undefined && first > 0 ? first : undefined,
		'w:hanging': first !== undefined && first < 0 ? -first : undefined,
		// A first line of exactly zero overrides an inherited hanging or first-line indent.
		...(first === 0 ? { 'w:firstLine': 0 } : {}),
	});
}
