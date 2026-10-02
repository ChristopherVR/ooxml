// Home > Styles: Conditional Formatting (quick rules, data bars, colour scales, icon sets, new,
// clear and manage), Format as Table and Cell Styles.
import type { ConditionalRule } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import {
	COLOR_SCALES,
	DATA_BAR_COLORS,
	ICON_SETS,
	colorScaleRule,
	dataBarRule,
	iconSetRule,
} from './cf-presets.js';
import { icon } from './icons.js';
import { TABLE_STYLES, cellStyleItems, tableStylePreview } from './style-presets.js';
import { editing, tableAt, target } from './util.js';

/** Quick rule kinds opened by the Highlight Cells and Top / Bottom menus. */
export const QUICK_RULES: ReadonlyArray<readonly [kind: string, label: string]> = [
	['greaterThan', 'Greater Than...'],
	['lessThan', 'Less Than...'],
	['between', 'Between...'],
	['equal', 'Equal To...'],
	['containsText', 'Text that Contains...'],
	['timePeriod', 'A Date Occurring...'],
	['duplicateValues', 'Duplicate Values...'],
	['top10', 'Top 10 Items...'],
	['top10Percent', 'Top 10%...'],
	['bottom10', 'Bottom 10 Items...'],
	['bottom10Percent', 'Bottom 10%...'],
	['aboveAverage', 'Above Average...'],
	['belowAverage', 'Below Average...'],
];

function addRule(ctx: EditorContext, rule: ConditionalRule): void {
	const t = target(ctx);
	if (!t) return;
	t.session.addConditionalFormat(t.sheet, { ranges: t.ranges, rules: [rule] });
}

export function styleCommands(): Command[] {
	return [
		editing({
			id: 'home.cf-quick',
			label: 'Highlight Cells Rules',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx, arg) =>
				void ctx.dialogs.open('cf-quick', { kind: typeof arg === 'string' ? arg : 'greaterThan' }),
		}),
		editing({
			id: 'home.cf-data-bar',
			label: 'Data Bars',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const preset = DATA_BAR_COLORS.find(([id]) => id === arg) ?? DATA_BAR_COLORS[0];
				if (preset) addRule(ctx, dataBarRule(preset[2]));
			},
		}),
		editing({
			id: 'home.cf-color-scale',
			label: 'Color Scales',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const preset = COLOR_SCALES.find(([id]) => id === arg) ?? COLOR_SCALES[0];
				if (preset) addRule(ctx, colorScaleRule(preset[2]));
			},
		}),
		editing({
			id: 'home.cf-icon-set',
			label: 'Icon Sets',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx, arg) => addRule(ctx, iconSetRule(typeof arg === 'string' ? arg : '3Arrows')),
		}),
		editing({
			id: 'home.cf-new-rule',
			label: 'New Rule...',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx) => void ctx.dialogs.open('cf-rule'),
		}),
		editing({
			id: 'home.cf-clear-selection',
			label: 'Clear Rules from Selected Cells',
			icon: icon('clear'),
			lock: 'formatCells',
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				t.session.batch('Clear rules', () => {
					for (const range of t.ranges) t.session.clearConditionalFormats(t.sheet, range);
				});
			},
		}),
		editing({
			id: 'home.cf-clear-sheet',
			label: 'Clear Rules from Entire Sheet',
			icon: icon('clear'),
			lock: 'formatCells',
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.clearConditionalFormats(t.sheet);
			},
		}),
		editing({
			id: 'home.cf-manage',
			label: 'Manage Rules...',
			icon: icon('conditional'),
			lock: 'formatCells',
			run: (ctx) => void ctx.dialogs.open('cf-manager'),
		}),
		editing({
			id: 'home.format-as-table',
			label: 'Format as Table',
			icon: icon('formatTable'),
			lock: false,
			run: (ctx, arg) => {
				const styleName = typeof arg === 'string' ? arg : 'TableStyleMedium2';
				const t = target(ctx);
				if (!t) return;
				const table = tableAt(t.ws, t.active);
				if (table) return t.session.updateTable(t.sheet, table.name, { styleName });
				void ctx.dialogs.open('create-table', { styleName });
			},
		}),
		editing({
			id: 'home.cell-styles',
			label: 'Cell Styles',
			icon: icon('cellStyles'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const t = target(ctx);
				if (t && typeof arg === 'string') t.session.applyCellStyle(t.sheet, t.ranges, arg);
			},
		}),
	];
}

export const tableStyleGalleryItems = (ctx: EditorContext) =>
	TABLE_STYLES.map((info) => ({
		id: info.name,
		label: ctx.t(`{family} Style {n}`, { family: ctx.t(info.family), n: info.index }),
		preview: tableStylePreview(
			info,
			ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' },
		),
	}));

export const cellStyleGalleryItems = (ctx: EditorContext) =>
	cellStyleItems(ctx.workbook()).map((item) => ({ ...item, label: ctx.t(item.label) }));

export { COLOR_SCALES, DATA_BAR_COLORS, ICON_SETS };
