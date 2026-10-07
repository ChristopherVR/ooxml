import { NS, first, parseXml } from '../../xml/index.js';
import type { ConditionalRule, Workbook } from '../model.js';
import { boolAttr } from '../read/xml-util.js';
import { parseColor } from '../read/style-parts.js';
import { resolveColor } from './colors.js';

interface BarAppearance {
	positive: string;
	negative: string;
	border: boolean;
	positiveBorder: string;
	negativeBorder: string;
	gradient: boolean;
	showValue: boolean;
	direction: 'context' | 'leftToRight' | 'rightToLeft';
}

/** Resolve modeled and linked bar appearance once per rule, using the shared color resolver. */
export function dataBarAppearance(
	rule: Extract<ConditionalRule, { type: 'dataBar' }>,
	workbook: Workbook,
): BarAppearance {
	const node = rule.extensionXml
		? first(parseXml(rule.extensionXml).documentElement, 'dataBar', NS.x14)
		: undefined;
	const color = (name: string, fallback: string) =>
		resolveColor(
			node ? parseColor(first(node, name, NS.x14), undefined) : undefined,
			workbook.theme,
			fallback,
		) ?? fallback;
	const positive = color(
		'fillColor',
		resolveColor(rule.color, workbook.theme, '#638EC6') ?? '#638EC6',
	);
	const negative = boolAttr(node, 'negativeBarColorSameAsPositive', !node)
		? positive
		: color('negativeFillColor', '#FF0000');
	const border = boolAttr(node, 'border', false);
	const positiveBorder = color('borderColor', positive);
	const negativeBorder = boolAttr(node, 'negativeBarBorderColorSameAsPositive', true)
		? positiveBorder
		: color('negativeBorderColor', negative);
	const direction = node?.getAttribute('direction');
	return {
		positive,
		negative,
		border,
		positiveBorder,
		negativeBorder,
		gradient: boolAttr(node, 'gradient', true),
		showValue: boolAttr(node, 'showValue', rule.showValue !== false),
		direction: direction === 'rightToLeft' || direction === 'leftToRight' ? direction : 'context',
	};
}
