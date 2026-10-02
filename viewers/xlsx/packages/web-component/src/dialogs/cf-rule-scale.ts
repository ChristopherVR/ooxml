// "Format all cells based on their values": 2- and 3-color scales, data bars and icon sets, with
// threshold types (lowest / highest value, number, percent, percentile, formula) and colours.
import type { CfvoThreshold, Color, ConditionalRule } from '@christophervr/xlsx-core';
import { ICON_SETS, iconCount, iconThresholds } from '../commands/cf-presets.js';
import type { EditorContext } from '../context.js';
import type { PanelRead, RulePanel } from './cf-rule-panels.js';
import { checkbox, el, field, row, select, textInput } from './fields.js';

type Style = '2color' | '3color' | 'dataBar' | 'iconSet';
type CfvoType = CfvoThreshold['type'];

const TYPES: ReadonlyArray<readonly [CfvoType, string]> = [
	['min', 'Lowest Value'],
	['max', 'Highest Value'],
	['num', 'Number'],
	['percent', 'Percent'],
	['percentile', 'Percentile'],
	['formula', 'Formula'],
];

const hex = (c: Color | undefined, fallback: string): string =>
	`#${(c?.rgb ?? fallback).slice(-6)}`;
const argb = (css: string): Color => ({ rgb: `FF${css.replace('#', '').toUpperCase()}` });

interface Point {
	element: HTMLElement;
	type: HTMLSelectElement;
	value: HTMLInputElement;
	color: HTMLInputElement;
}

function point(
	ctx: EditorContext,
	label: string,
	types: CfvoType[],
	initial: CfvoType,
	color?: string,
): Point {
	const type = select(
		ctx,
		TYPES.filter(([t]) => types.includes(t)),
		initial,
	);
	const value = textInput(ctx);
	const pick = el(ctx, 'input', 'xve-input');
	pick.type = 'color';
	pick.value = color ?? '#000000';
	const fields: HTMLElement[] = [
		field(ctx, `${label} type`, type),
		field(ctx, `${label} value`, value),
	];
	if (color) fields.push(field(ctx, `${label} color`, pick));
	const sync = (): void => {
		value.disabled = type.value === 'min' || type.value === 'max';
	};
	type.addEventListener('change', sync);
	sync();
	return { element: row(ctx, ...fields), type, value, color: pick };
}

const readPoint = (p: Point): CfvoThreshold | { error: string; control: HTMLElement } => {
	const type = p.type.value as CfvoType;
	if (type === 'min' || type === 'max') return { type };
	const raw = p.value.value.trim().replace(/^=/, '');
	if (!raw) return { error: 'Enter a value.', control: p.value };
	if (type !== 'formula' && !Number.isFinite(Number(raw)))
		return { error: 'Enter a valid number.', control: p.value };
	return { type, value: raw };
};

const loadPoint = (p: Point, cfvo: CfvoThreshold | undefined, color?: Color): void => {
	if (cfvo) {
		p.type.value = cfvo.type;
		p.value.value = cfvo.value ?? '';
		p.type.dispatchEvent(new Event('change'));
	}
	if (color) p.color.value = hex(color, '000000');
};

export function valuesPanel(ctx: EditorContext): RulePanel {
	const style = select(
		ctx,
		[
			['2color', '2-Color Scale'],
			['3color', '3-Color Scale'],
			['dataBar', 'Data Bar'],
			['iconSet', 'Icon Set'],
		],
		'2color',
	);
	const all: CfvoType[] = ['min', 'max', 'num', 'percent', 'percentile', 'formula'];
	const min = point(
		ctx,
		'Minimum',
		['min', 'num', 'percent', 'percentile', 'formula'],
		'min',
		'#F8696B',
	);
	const mid = point(
		ctx,
		'Midpoint',
		['num', 'percent', 'percentile', 'formula'],
		'percentile',
		'#FFEB84',
	);
	mid.value.value = '50';
	const max = point(
		ctx,
		'Maximum',
		['max', 'num', 'percent', 'percentile', 'formula'],
		'max',
		'#63BE7B',
	);
	const bar = el(ctx, 'input', 'xve-input');
	bar.type = 'color';
	bar.value = '#638EC6';
	const barField = field(ctx, 'Bar color', bar);
	const iconSet = select(ctx, ICON_SETS, '3Arrows');
	const reverse = checkbox(ctx, 'Reverse Icon Order');
	const iconOnly = checkbox(ctx, 'Show Icon Only');
	const iconRows = el(ctx, 'div', 'xve-cf-icons');
	let icons: Point[] = [];
	const buildIcons = (): void => {
		const n = iconCount(iconSet.value);
		const defaults = iconThresholds(n);
		icons = defaults.slice(1).map((cfvo, i) => {
			const p = point(
				ctx,
				`Icon ${i + 2}`,
				all.filter((t) => t !== 'min' && t !== 'max'),
				'percent',
			);
			p.value.value = cfvo.value ?? '';
			return p;
		});
		iconRows.replaceChildren(...icons.map((p) => p.element));
	};
	iconSet.addEventListener('change', buildIcons);
	buildIcons();
	const iconFields = el(ctx, 'div');
	iconFields.append(field(ctx, 'Icon Style', iconSet), reverse.wrapper, iconOnly.wrapper, iconRows);
	const sync = (): void => {
		const v = style.value as Style;
		min.element.hidden = v === 'iconSet';
		max.element.hidden = v === 'iconSet';
		mid.element.hidden = v !== '3color';
		barField.hidden = v !== 'dataBar';
		iconFields.hidden = v !== 'iconSet';
		for (const p of [min, max]) {
			const colorField = p.color.closest('label');
			if (colorField) colorField.hidden = v === 'dataBar';
		}
	};
	style.addEventListener('change', sync);
	sync();
	const element = el(ctx, 'div', 'xve-cf-panel');
	element.append(
		field(ctx, 'Format Style', style),
		min.element,
		mid.element,
		max.element,
		barField,
		iconFields,
	);
	return {
		element,
		styled: false,
		load(rule) {
			if (rule.type === 'colorScale') {
				const three = rule.colors.length >= 3;
				style.value = three ? '3color' : '2color';
				loadPoint(min, rule.thresholds[0], rule.colors[0]);
				if (three) loadPoint(mid, rule.thresholds[1], rule.colors[1]);
				loadPoint(max, rule.thresholds[three ? 2 : 1], rule.colors[three ? 2 : 1]);
			} else if (rule.type === 'dataBar') {
				style.value = 'dataBar';
				loadPoint(min, rule.min);
				loadPoint(max, rule.max);
				bar.value = hex(rule.color, '638EC6');
			} else if (rule.type === 'iconSet') {
				style.value = 'iconSet';
				iconSet.value = rule.iconSet;
				buildIcons();
				rule.thresholds.slice(1).forEach((cfvo, i) => {
					const p = icons[i];
					if (p) loadPoint(p, cfvo);
				});
				reverse.input.checked = !!rule.reverse;
				iconOnly.input.checked = rule.showValue === false;
			}
			sync();
		},
		read(): PanelRead {
			const v = style.value as Style;
			if (v === 'iconSet') {
				const thresholds: CfvoThreshold[] = [{ type: 'percent', value: '0' }];
				for (const p of icons) {
					const r = readPoint(p);
					if ('error' in r) return r;
					thresholds.push(r);
				}
				const rule: ConditionalRule = {
					type: 'iconSet',
					iconSet: iconSet.value,
					thresholds,
					priority: 1,
				};
				if (reverse.input.checked) rule.reverse = true;
				if (iconOnly.input.checked) rule.showValue = false;
				return rule;
			}
			const points = v === '3color' ? [min, mid, max] : [min, max];
			const thresholds: CfvoThreshold[] = [];
			for (const p of points) {
				const r = readPoint(p);
				if ('error' in r) return r;
				thresholds.push(r);
			}
			if (v === 'dataBar')
				return {
					type: 'dataBar',
					min: thresholds[0]!,
					max: thresholds[1]!,
					color: argb(bar.value),
					priority: 1,
				};
			return {
				type: 'colorScale',
				thresholds,
				colors: points.map((p) => argb(p.color.value)),
				priority: 1,
			};
		},
		focus: () => style.focus(),
	};
}
