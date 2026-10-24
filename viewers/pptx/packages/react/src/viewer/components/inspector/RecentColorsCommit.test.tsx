// @vitest-environment happy-dom
/**
 * Every native `<input type="color">` in the inspector adds to the deck's
 * "Recent Colors" list (`p:clrMru`) once per COMMITTED pick, never while the
 * user drags through the OS picker.
 *
 * A native colour input fires `input` continuously during a drag and `change`
 * once when the pick settles. React's `onChange` runs on every `input`, so a
 * push from `onChange` filled the capped list with every shade passed on the
 * way. Each case below drags through a few shades (only `input` events), checks
 * that the live apply still ran but nothing was pushed, then commits (`change`)
 * and checks that exactly the final colour was pushed, once.
 */
import type {
	PptxChartAxisFormatting,
	PptxChartSeries,
	PptxElement,
	PptxTableCellStyle,
	TextStyle,
} from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { Mock } from 'vitest';

import { AfterAnimationRow } from './AfterAnimationRow';
import { ChartAxisStyleOptions } from './ChartAxisStyleOptions';
import { ChartDataPointMarkerOptions } from './ChartDataPointMarkerOptions';
import { ChartDataPointOptions } from './ChartDataPointOptions';
import { ChartMarkerOptions } from './ChartMarkerOptions';
import { DebouncedColorInput } from './DebouncedColorInput';
import { DuotonePanel } from './DuotonePanel';
import { ColorPickerRow, GradientStopRow } from './FillStrokeSubComponents';
import { InspectorFields } from './InspectorFields';
import { Text3DProperties } from './properties/Text3DProperties';
import { RecentColorsProvider } from './RecentColorsContext';
import { TableCellAdvancedFill } from './TableCellAdvancedFill';
import { GradientControls, PatternControls } from './TableCellFillSubControls';
import { TableCellPatternFill } from './TableCellPatternFill';
import { TextEffectsPanel } from './TextEffectsPanel';
import { AdvancedTextFormatting } from './TextPropertiesHelpers';

vi.mock(import('../../i18n/use-translation'), async (importOriginal) => ({
	...(await importOriginal()),
	useTranslation: (() => ({ t: (key: string) => key })) as never,
}));

/** Shades a drag passes through, then the colour the picker settles on. */
const DRAG = ['#110000', '#220000', '#330000'];
const FINAL = '#440000';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

/** Write `value` the way the browser does, so React's value tracker sees a change. */
function setValue(input: HTMLInputElement, value: string): void {
	Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set?.call(input, value);
}

function drag(input: HTMLInputElement): void {
	for (const shade of DRAG) {
		act(() => {
			setValue(input, shade);
			input.dispatchEvent(new Event('input', { bubbles: true }));
		});
	}
}

/**
 * The picker settles on `FINAL`. The harness props never change, so React has
 * put each input back to its prop value after every drag tick; a real host
 * re-renders with the dragged colour, so the value is set again first.
 */
function settle(input: HTMLInputElement): void {
	act(() => {
		setValue(input, FINAL);
		input.dispatchEvent(new Event('change', { bubbles: true }));
	});
}

const noNum = () => () => {};

const stops = [
	{ color: '#ff0000', position: 0 },
	{ color: '#0000ff', position: 100 },
];

const lineSeries: PptxChartSeries[] = [
	{
		name: 'S1',
		values: [1],
		marker: { symbol: 'circle' },
		dataPoints: [{ idx: 0, marker: { symbol: 'circle' } }],
	} as PptxChartSeries,
];

interface Case {
	name: string;
	/** How many native colour inputs the case renders. */
	inputs: number;
	/** False where the input only applies on commit (the duotone picker). */
	live?: boolean;
	render(apply: Mock): React.ReactElement;
}

const CASES: Case[] = [
	{
		name: 'AfterAnimationRow',
		inputs: 1,
		render: (apply) => (
			<AfterAnimationRow
				action='dimToColor'
				color='#000000'
				canEdit
				onActionChange={apply}
				onColorChange={apply}
			/>
		),
	},
	{
		name: 'ChartAxisStyleOptions',
		inputs: 3,
		render: (apply) => (
			<ChartAxisStyleOptions
				axes={
					[
						{ axisType: 'valAx', majorGridlines: true, minorGridlines: true },
					] as PptxChartAxisFormatting[]
				}
				canEdit
				onSetLogScale={apply}
				onSetTitleStyle={apply}
				onSetGridlineStyle={apply}
			/>
		),
	},
	{
		name: 'ChartDataPointMarkerOptions',
		inputs: 1,
		render: (apply) => (
			<ChartDataPointMarkerOptions
				chartType='line'
				categories={['A']}
				series={lineSeries}
				canEdit
				onSetPointMarker={apply}
			/>
		),
	},
	{
		name: 'ChartDataPointOptions',
		inputs: 1,
		render: (apply) => (
			<ChartDataPointOptions
				chartType='bar'
				categories={['A']}
				series={lineSeries}
				canEdit
				onSetPointFill={apply}
				onSetPointExplosion={apply}
				onSetPointLabel={apply}
			/>
		),
	},
	{
		name: 'ChartMarkerOptions',
		inputs: 1,
		render: (apply) => (
			<ChartMarkerOptions chartType='line' series={lineSeries} canEdit onSetMarker={apply} />
		),
	},
	{
		name: 'DebouncedColorInput',
		inputs: 1,
		render: (apply) => <DebouncedColorInput value='#000000' onCommit={apply} />,
	},
	{
		name: 'DuotonePanel',
		inputs: 2,
		live: false,
		render: (apply) => (
			<DuotonePanel
				selectedElement={{ type: 'image', id: 'i', imageEffects: {} } as unknown as PptxElement}
				canEdit
				onUpdateElement={apply}
			/>
		),
	},
	{
		// The generic renderer behind the effects, line, size and text box sections.
		name: 'InspectorFields',
		inputs: 1,
		render: (apply) => (
			<InspectorFields
				sections={[
					{
						id: 'shadow',
						fields: [
							{
								id: 'color',
								kind: 'color',
								labelKey: 'pptx.effects.color',
								read: () => '#000000',
								write: (_target: unknown, value) => ({ shadowColor: String(value) }),
							},
						],
					},
				]}
				target={{}}
				onPatch={apply}
			/>
		),
	},
	{
		name: 'ColorPickerRow',
		inputs: 1,
		render: (apply) => (
			<ColorPickerRow label='Fill' value='#000000' prefix='fill' onChange={apply} />
		),
	},
	{
		name: 'GradientStopRow',
		inputs: 1,
		render: (apply) => (
			<GradientStopRow stop={stops[0]} index={0} total={2} allStops={stops} onUpdate={apply} />
		),
	},
	...(['gradient', 'pattern'] as const).map((fillMode): Case => ({
		name: `TableCellAdvancedFill (${fillMode})`,
		inputs: 2,
		render: (apply) => (
			<TableCellAdvancedFill
				cellStyle={{ fillMode, gradientFillStops: stops } as PptxTableCellStyle}
				canEdit
				onUpdateCellStyle={apply}
			/>
		),
	})),
	{
		name: 'TableCellPatternFill',
		inputs: 2,
		render: (apply) => <TableCellPatternFill cellStyle={{}} canEdit onUpdateCellStyle={apply} />,
	},
	{
		name: 'TableCellFillSubControls GradientControls',
		inputs: 2,
		render: (apply) => (
			<GradientControls
				cellStyle={{ gradientFillStops: stops } as PptxTableCellStyle}
				canEdit
				onUpdateCellStyle={apply}
			/>
		),
	},
	{
		name: 'TableCellFillSubControls PatternControls',
		inputs: 2,
		render: (apply) => <PatternControls cellStyle={{}} canEdit onUpdateCellStyle={apply} />,
	},
	{
		name: 'TextEffectsPanel',
		inputs: 2,
		render: (apply) => (
			<TextEffectsPanel
				ts={{ textShadowColor: '#000000', textGlowColor: '#ffff00' } as TextStyle}
				onUpdateTextStyle={apply}
				numChange={noNum}
			/>
		),
	},
	{
		name: 'AdvancedTextFormatting',
		inputs: 3,
		render: (apply) => (
			<AdvancedTextFormatting
				ts={{ textOutlineWidth: 1 } as TextStyle}
				canEdit
				onUpdateTextStyle={apply}
				numChange={noNum}
			/>
		),
	},
	{
		name: 'Text3DProperties',
		inputs: 1,
		render: (apply) => (
			<Text3DProperties
				ts={{ text3d: { extrusionHeight: 12700 } } as TextStyle}
				onUpdateTextStyle={apply}
			/>
		),
	},
];

describe('recent colours grow once per committed pick, not per drag tick', () => {
	it.each(CASES)('$name', ({ inputs: count, live = true, render }) => {
		const apply = vi.fn();
		const pushColor = vi.fn();
		act(() => {
			root.render(
				<RecentColorsProvider value={{ recentColors: [], pushColor }}>
					{render(apply)}
				</RecentColorsProvider>,
			);
		});
		const inputs = Array.from(container.querySelectorAll<HTMLInputElement>('input[type="color"]'));
		expect(inputs).toHaveLength(count);

		for (const [index, input] of inputs.entries()) {
			apply.mockClear();
			pushColor.mockClear();

			drag(input);
			if (live) {
				expect(apply, `input ${index} applies live while dragging`).toHaveBeenCalledTimes(
					DRAG.length,
				);
			}
			expect(pushColor, `input ${index} pushes nothing while dragging`).not.toHaveBeenCalled();

			settle(input);
			expect(
				pushColor,
				`input ${index} pushes the settled colour once`,
			).toHaveBeenCalledExactlyOnceWith(FINAL);
		}
	});
});

describe('a picker focused and left without a pick changes nothing', () => {
	it.each(CASES)('$name', ({ inputs: count, render }) => {
		const apply = vi.fn();
		const pushColor = vi.fn();
		act(() => {
			root.render(
				<RecentColorsProvider value={{ recentColors: [], pushColor }}>
					{render(apply)}
				</RecentColorsProvider>,
			);
		});
		const inputs = Array.from(container.querySelectorAll<HTMLInputElement>('input[type="color"]'));
		expect(inputs).toHaveLength(count);

		for (const [index, input] of inputs.entries()) {
			act(() => {
				input.focus();
				input.blur();
			});
			expect(apply, `input ${index} applies nothing`).not.toHaveBeenCalled();
			expect(pushColor, `input ${index} adds no recent colour`).not.toHaveBeenCalled();
		}
	});
});
