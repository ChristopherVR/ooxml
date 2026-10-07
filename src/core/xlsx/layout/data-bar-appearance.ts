import { NS, first, parseXml, elements } from '../../xml/index.js';
import type { ConditionalRule, Workbook } from '../model.js';
import { boolAttr } from '../read/xml-util.js';
import { dataBarLengths } from '../conditional-extensions.js';
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
	axis: 'automatic' | 'middle' | 'none';
	axisColor: string;
	autoMin: boolean;
	autoMax: boolean;
	minLength: number;
	maxLength: number;
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
	const axis = node?.getAttribute('axisPosition');
	const lengths = dataBarLengths(rule, node);
	const thresholds = node
		? elements(node).filter((n) => n.localName === 'cfvo' && n.namespaceURI === NS.x14)
		: [];
	return {
		positive,
		negative,
		border,
		positiveBorder,
		negativeBorder,
		gradient: boolAttr(node, 'gradient', true),
		showValue: boolAttr(node, 'showValue', rule.showValue !== false),
		direction: direction === 'rightToLeft' || direction === 'leftToRight' ? direction : 'context',
		axis: node ? (axis === 'middle' || axis === 'none' ? axis : 'automatic') : 'none',
		axisColor: color('axisColor', '#000000'),
		autoMin: rule.min.type === 'min' && thresholds[0]?.getAttribute('type') === 'autoMin',
		autoMax: rule.max.type === 'max' && thresholds[1]?.getAttribute('type') === 'autoMax',
		minLength: lengths.minLength / 100,
		maxLength: lengths.maxLength / 100,
	};
}
