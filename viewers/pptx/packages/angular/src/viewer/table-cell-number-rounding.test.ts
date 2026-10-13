/**
 * The table inspector's number inputs show two decimals at most.
 *
 * The value helpers the templates bind to are called directly (no TestBed,
 * see `vitest.config.ts`), and the bindings that reach them are pinned as a
 * source-text guard, like the other component tests in this package.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { PptxTableCellStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';
import { TableCellAdvancedFillComponent } from './table-cell-advanced-fill.component';
import { TableCellFormattingComponent } from './table-cell-formatting.component';
import { TablePropertiesComponent } from './table-properties.component';

// 1.52pt borders, 4pt margins and a 29.2pt row, read back from EMU as px.
const BORDER = 19304 / 9525;
const MARGIN = 50800 / 9525;
const ROW = 370840 / 9525;

const STYLE: PptxTableCellStyle = {
	fontSize: 10.5,
	borderTopWidth: BORDER,
	marginLeft: MARGIN,
	gradientFillAngle: 33.333333,
};

function create<T>(make: () => T): T {
	return runInInjectionContext(Injector.create({ providers: [] }), make);
}

interface FormattingValues {
	style: () => PptxTableCellStyle;
	widthOf: (key: 'borderTopWidth') => number;
	fontSizeValue: () => number;
}

interface FillValues {
	cellStyle: () => PptxTableCellStyle;
	marginValue: (key: 'marginLeft') => number;
	gradientAngleValue: () => number;
}

describe('table inspector number display', () => {
	it('rounds the cell font size and border width', () => {
		const component = create(
			() => new TableCellFormattingComponent(),
		) as unknown as FormattingValues;
		Object.defineProperty(component, 'style', { value: () => STYLE });

		expect(component.widthOf('borderTopWidth')).toBe(2.03);
		expect(component.fontSizeValue()).toBe(10.5);
	});

	it('rounds the cell margins and gradient angle', () => {
		const component = create(() => new TableCellAdvancedFillComponent()) as unknown as FillValues;
		Object.defineProperty(component, 'cellStyle', { value: () => STYLE });

		expect(component.marginValue('marginLeft')).toBe(5.33);
		expect(component.gradientAngleValue()).toBe(33.33);
	});

	it('rounds the row height', () => {
		const component = create(() => new TablePropertiesComponent()) as unknown as {
			rowHeightValue: (height: number | undefined) => number;
		};

		expect(component.rowHeightValue(ROW)).toBe(38.93);
	});

	it('binds the inputs to the rounded values', () => {
		const formatting = componentSource(import.meta.dirname, 'table-cell-formatting.component.ts');
		const fill = componentSource(import.meta.dirname, 'table-cell-advanced-fill.component.ts');
		const properties = componentSource(import.meta.dirname, 'table-properties.component.ts');
		expect(formatting).toContain('[value]="fontSizeValue()"');
		expect(formatting).toContain('[value]="widthOf(edge.widthKey)"');
		expect(fill).toContain('[value]="gradientAngleValue()"');
		expect(properties).toContain('[value]="rowHeightValue(row.height)"');
	});
});
