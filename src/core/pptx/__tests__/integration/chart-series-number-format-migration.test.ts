import { classicChartDecks } from './chart-save-parity';
import { describeFamilyMigration, type FamilyEdit } from './chart-family-migration';

/**
 * A series' own number format (`chart-neutral-save-series-format.ts`) saved through the neutral
 * writer and the object-tree writers over every committed deck. The object-tree writers never
 * wrote the field, so the edit was lost on save; the neutral save writes it as the series' label
 * number format.
 */

const DECKS = await classicChartDecks();

const EDITS: FamilyEdit[] = [
	{
		name: 'series number format',
		applies: (chart) => chart.series.length > 0,
		edit: (chart) => {
			const series = chart.series[0]!;
			series.numberFormat = series.numberFormat === '0.0%' ? '#,##0.00' : '0.0%';
		},
		read: (chart) => chart.series[0]!.numberFormat,
	},
];

/**
 * Apart from the series `c:dLbls`, the object-tree parts differ in what those writers do to the rest
 * of an edited part (theme title font pinning, JavaScript number caches, line series `c:idx`,
 * dropped switched-off pie `c:dLbls`), as recorded in `chart-group-flags-migration.test.ts`.
 */
describeFamilyMigration(DECKS, EDITS, {
	'series number format': { parts: 75, legacyKept: 0, xml: 0 },
});
