import { classicChartDecks } from './chart-save-parity';
import { describeFamilyMigration, type FamilyEdit } from './chart-family-migration';

/**
 * The chart-wide gridline flag (`chart-neutral-save-gridlines.ts`) saved through the neutral
 * writer and the object-tree writers over every committed deck. The object-tree writers never
 * wrote the flag, so a flag edited on its own was lost on save; the neutral save sets the major
 * gridlines of the first value axis, where the flag is read from.
 */

const DECKS = await classicChartDecks();

const EDITS: FamilyEdit[] = [
	{
		name: 'gridline flag',
		applies: (chart) => chart.axes?.some((axis) => axis.axisType === 'valAx') === true,
		edit: (chart) => {
			chart.style = { ...chart.style, hasGridlines: !chart.style?.hasGridlines };
		},
		// A reloaded chart does not always carry the flag; the axis it is read from decides.
		read: (chart) =>
			chart.style?.hasGridlines ??
			chart.axes?.find((axis) => axis.axisType === 'valAx')?.majorGridlines ??
			false,
	},
];

/**
 * The object-tree save kept the flag only on the 5 parts whose value axis already drew gridlines
 * the flag had not reported (the edit asked for what was there). Apart from the gridlines, the object-tree parts differ in what those writers do to the rest of an
 * edited part (theme title font pinning, JavaScript number caches, line series `c:idx`), as
 * recorded in `chart-group-flags-migration.test.ts`.
 */
describeFamilyMigration(DECKS, EDITS, {
	'gridline flag': { parts: 51, legacyKept: 5, xml: 4 },
});
