import {
	formatRange,
	parseRange,
	resolveColor,
	selectedSparklineGroups,
	type SparklineGroupPatch,
	type SparklineType,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, fieldset, numberInput, select, text, textInput } from './fields';
import { showDialog } from './frame';

const TYPES = [
	['line', 'Line'],
	['column', 'Column'],
	['stacked', 'Win/Loss'],
] as const;

export function openCreateSparklines(
	ctx: EditorContext,
	type: SparklineType = 'line',
): Promise<boolean | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const data = textInput(ctx);
	const locations = textInput(ctx, formatRange(t.range));
	const kind = select(ctx, TYPES, type);
	const dates = textInput(ctx);
	return showDialog(ctx, {
		name: 'create-sparklines',
		heading: 'Create Sparklines',
		body: [
			field(ctx, 'Data Range:', data),
			field(ctx, 'Location Range:', locations),
			field(ctx, 'Sparkline Type:', kind),
			field(ctx, 'Date Range (optional):', dates),
		],
		opened: () => data.focus(),
		submit: () => {
			const range = parseRange(locations.value.trim().replace(/^=/, '').replaceAll('$', ''));
			if (!range) throw new Error('The reference is not valid.');
			t.session.batch('Create sparklines', () => {
				t.session.createSparklineGroup(t.sheet, data.value, range, {
					type: kind.value as SparklineType,
				});
				if (dates.value.trim()) t.session.setSparklineDateRange(t.sheet, [range], dates.value);
			});
			ctx.selection.set({ ranges: [range], active: range.start, anchor: range.start });
			return true;
		},
	});
}

export function openSparklineSettings(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const groups = t ? selectedSparklineGroups(t.ws, t.ranges) : [];
	const group = groups[0];
	if (!t || !group) return Promise.resolve(undefined);
	const dates = textInput(ctx, group.dateAxis ? (group.dateFormula ?? '') : '');
	if (
		groups.some(
			(entry) => entry.dateAxis !== group.dateAxis || entry.dateFormula !== group.dateFormula,
		)
	) {
		dates.value = '';
		dates.placeholder = ctx.t('Mixed');
	}
	let datesChanged = false;
	dates.addEventListener('input', () => {
		datesChanged = true;
	});
	dates.addEventListener('change', () => {
		datesChanged = true;
	});
	const changes = new Map<keyof SparklineGroupPatch, () => SparklineGroupPatch>();
	const track = (
		input: HTMLInputElement | HTMLSelectElement,
		key: keyof SparklineGroupPatch,
		read: () => SparklineGroupPatch,
	) => {
		const changed = () => {
			if (input.dataset.mixed && input.value === '') changes.delete(key);
			else changes.set(key, read);
		};
		input.addEventListener('input', changed);
		input.addEventListener('change', changed);
	};
	const mixed = (input: HTMLInputElement | HTMLSelectElement, key: keyof SparklineGroupPatch) => {
		if (!groups.some((entry) => entry[key] !== group[key])) return;
		input.dataset.mixed = 'true';
		if (input instanceof ctx.host.ownerDocument.defaultView!.HTMLSelectElement) {
			const option = ctx.host.ownerDocument.createElement('option');
			option.value = '';
			option.textContent = ctx.t('Mixed');
			input.prepend(option);
			input.value = '';
		} else if (input.type === 'checkbox') {
			input.indeterminate = true;
			input.checked = false;
		} else {
			input.value = '';
			input.placeholder = ctx.t('Mixed');
		}
	};
	const kind = select(ctx, TYPES, group.type);
	const width = numberInput(ctx, group.lineWeight, 0, 1584, 0.25);
	const empty = select(
		ctx,
		[
			['gap', 'Gaps'],
			['zero', 'Zero'],
			['span', 'Connect data points'],
		],
		group.displayEmptyCellsAs,
	);
	const markers = [
		['markers', 'Markers'],
		['high', 'High Point'],
		['low', 'Low Point'],
		['first', 'First Point'],
		['last', 'Last Point'],
		['negative', 'Negative Points'],
	] as const;
	const checks = markers.map(([key, label]) => ({ key, ...checkbox(ctx, label, group[key]) }));
	const hidden = checkbox(ctx, 'Show data in hidden rows and columns', group.displayHidden);
	const axis = checkbox(ctx, 'Show horizontal axis', group.displayXAxis);
	const rtl = checkbox(ctx, 'Plot right to left', group.rightToLeft);
	const scales = [
		['individual', 'Individual'],
		['group', 'Same for all sparklines'],
		['custom', 'Custom'],
	] as const;
	const minType = select(ctx, scales, group.minAxisType);
	const maxType = select(ctx, scales, group.maxAxisType);
	const min = numberInput(ctx, group.manualMin ?? 0, -1e308, 1e308, 0.1);
	const max = numberInput(ctx, group.manualMax ?? 1, -1e308, 1e308, 0.1);
	const updateLimits = () => {
		min.disabled = minType.value !== 'custom';
		max.disabled = maxType.value !== 'custom';
	};
	minType.addEventListener('change', updateLimits);
	maxType.addEventListener('change', updateLimits);
	updateLimits();
	const color = textInput(ctx, resolveColor(group.colorSeries, t.workbook.theme));
	color.type = 'color';
	const negative = textInput(ctx, resolveColor(group.colorNegative, t.workbook.theme));
	negative.type = 'color';
	const mixedColors = groups.some(
		(entry) =>
			(resolveColor(entry.colorSeries, t.workbook.theme) ?? '#000000').toLowerCase() !==
				color.value.toLowerCase() ||
			(resolveColor(entry.colorNegative, t.workbook.theme) ?? '#000000').toLowerCase() !==
				negative.value.toLowerCase(),
	);
	for (const [input, key] of [
		[kind, 'type'],
		[width, 'lineWeight'],
		[empty, 'displayEmptyCellsAs'],
		[hidden.input, 'displayHidden'],
		[axis.input, 'displayXAxis'],
		[rtl.input, 'rightToLeft'],
		[minType, 'minAxisType'],
		[maxType, 'maxAxisType'],
		[min, 'manualMin'],
		[max, 'manualMax'],
	] as const)
		mixed(input, key);
	for (const check of checks) {
		mixed(check.input, check.key);
		track(check.input, check.key, () => ({ [check.key]: check.input.checked }));
	}
	track(kind, 'type', () => ({ type: kind.value as SparklineType }));
	track(width, 'lineWeight', () => ({ lineWeight: Number(width.value) }));
	track(empty, 'displayEmptyCellsAs', () => ({
		displayEmptyCellsAs: empty.value as 'gap' | 'zero' | 'span',
	}));
	track(hidden.input, 'displayHidden', () => ({ displayHidden: hidden.input.checked }));
	track(axis.input, 'displayXAxis', () => ({ displayXAxis: axis.input.checked }));
	track(rtl.input, 'rightToLeft', () => ({ rightToLeft: rtl.input.checked }));
	track(minType, 'minAxisType', () => ({
		minAxisType: minType.value as 'individual' | 'group' | 'custom',
		...(minType.value === 'custom' ? { manualMin: Number(min.value) } : {}),
	}));
	track(maxType, 'maxAxisType', () => ({
		maxAxisType: maxType.value as 'individual' | 'group' | 'custom',
		...(maxType.value === 'custom' ? { manualMax: Number(max.value) } : {}),
	}));
	track(min, 'manualMin', () => ({ manualMin: Number(min.value) }));
	track(max, 'manualMax', () => ({ manualMax: Number(max.value) }));
	track(color, 'colorSeries', () => ({ colorSeries: { rgb: color.value.slice(1) } }));
	track(negative, 'colorNegative', () => ({ colorNegative: { rgb: negative.value.slice(1) } }));
	updateLimits();
	return showDialog(ctx, {
		name: 'sparkline-settings',
		heading: 'Sparkline Settings',
		wide: true,
		body: [
			text(ctx, 'Changes apply to the entire selected sparkline groups.'),
			text(ctx, 'Only changed settings are applied.'),
			field(ctx, 'Sparkline Type:', kind),
			field(ctx, 'Line Weight:', width),
			field(ctx, 'Date Range (optional):', dates),
			text(ctx, 'Leave the date range empty for a general axis.'),
			fieldset(ctx, 'Show', ...checks.map((check) => check.wrapper)),
			fieldset(
				ctx,
				'Colors',
				field(ctx, 'Sparkline Color:', color),
				field(ctx, 'Negative Color:', negative),
				...(mixedColors ? [text(ctx, 'Colors vary between selected groups.')] : []),
			),
			fieldset(
				ctx,
				'Axis',
				field(ctx, 'Minimum Axis:', minType),
				field(ctx, 'Minimum Value:', min),
				field(ctx, 'Maximum Axis:', maxType),
				field(ctx, 'Maximum Value:', max),
				axis.wrapper,
				rtl.wrapper,
			),
			field(ctx, 'Empty Cells:', empty),
			hidden.wrapper,
		],
		submit: () => {
			const patch: SparklineGroupPatch = {};
			for (const read of changes.values()) Object.assign(patch, read());
			if (minType.value !== 'custom') delete patch.manualMin;
			if (maxType.value !== 'custom') delete patch.manualMax;
			if (
				(patch.minAxisType === 'custom' && !min.value.trim()) ||
				(patch.maxAxisType === 'custom' && !max.value.trim())
			)
				throw new Error('The sparkline axis limits are not valid.');
			t.session.batch('Format sparklines', () => {
				if (datesChanged) t.session.setSparklineDateRange(t.sheet, t.ranges, dates.value);
				if (Object.keys(patch).length) t.session.setSparklineGroups(t.sheet, t.ranges, patch);
			});
			return true;
		},
	});
}
