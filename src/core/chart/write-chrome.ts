// `c:chartSpace` writer: the chrome beside the plot, in the sequences of ECMA-376 Part 1, 21.2:
// floor and walls (CT_Surface: thickness, spPr, pictureOptions, extLst), the data table (CT_DTable:
// showHorzBorder, showVertBorder, showOutline, showKeys, spPr, txPr, extLst) and display units
// (CT_DispUnits: custUnit or builtInUnit, dispUnitsLbl with layout, tx, spPr, txPr; extLst).
import type { ChartDataTable, ChartDisplayUnits, ChartSurface } from './model-chrome';
import { layoutXml, shapePropertiesXml, textBodyXml } from './write-shape';
import { chartTextXml } from './write-text';
import { elementXml, raw, valXml, type ChartWriteContext } from './write-util';

/** A floor or wall (`local` is `floor`, `sideWall` or `backWall`). */
export function surfaceXml(
	context: ChartWriteContext,
	local: string,
	surface: ChartSurface | undefined,
): string {
	if (!surface) return '';
	return elementXml(
		local,
		valXml('thickness', surface.thickness) +
			shapePropertiesXml(context, surface.spPr) +
			raw(context, surface.pictureOptionsXml) +
			raw(context, surface.extLst),
	);
}

/** The data table (`c:dTable`). */
export function dataTableXml(
	context: ChartWriteContext,
	table: ChartDataTable | undefined,
): string {
	if (!table) return '';
	return elementXml(
		'dTable',
		valXml('showHorzBorder', table.showHorizontalBorder) +
			valXml('showVertBorder', table.showVerticalBorder) +
			valXml('showOutline', table.showOutline) +
			valXml('showKeys', table.showKeys) +
			shapePropertiesXml(context, table.spPr) +
			textBodyXml(context, table.txPr) +
			raw(context, table.extLst),
	);
}

/** Display units of a value axis (`c:dispUnits`). */
export function displayUnitsXml(
	context: ChartWriteContext,
	units: ChartDisplayUnits | undefined,
): string {
	if (!units) return '';
	const label = units.label;
	return elementXml(
		'dispUnits',
		valXml('custUnit', units.customUnit) +
			(units.customUnit === undefined ? valXml('builtInUnit', units.builtInUnit) : '') +
			(label
				? elementXml(
						'dispUnitsLbl',
						layoutXml(context, label.layout) +
							chartTextXml(context, label.tx, false) +
							shapePropertiesXml(context, label.spPr) +
							textBodyXml(context, label.txPr),
					)
				: '') +
			raw(context, units.extLst),
	);
}
