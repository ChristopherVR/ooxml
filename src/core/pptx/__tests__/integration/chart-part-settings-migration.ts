import { describeFamilyMigrationEdits, type FamilyEdit } from './chart-family-migration';

/**
 * Chart part settings (`chart-neutral-save-part-settings.ts`) saved through the neutral writer and
 * the object-tree writers over every committed deck. No committed deck holds chart print settings
 * or pivot formats: those are covered by `chart-neutral-save-part-settings.test.ts` and the
 * round-trip suites (`chart-print-settings-roundtrip.test.ts`).
 */

const MAPPING = {
	bg1: 'dk1',
	tx1: 'lt1',
	bg2: 'dk2',
	tx2: 'lt2',
	accent1: 'accent2',
	accent2: 'accent1',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hlink: 'hlink',
	folHlink: 'folHlink',
};

const EDITS: FamilyEdit[] = [
	{
		name: 'external data auto-update',
		applies: (chart) => chart.externalData !== undefined,
		edit: (chart) => {
			chart.externalData = { ...chart.externalData!, autoUpdate: !chart.externalData!.autoUpdate };
		},
		read: (chart) => chart.externalData?.autoUpdate ?? false,
	},
	{
		name: 'colour map override',
		applies: () => true,
		edit: (chart) => {
			chart.clrMapOvr = chart.clrMapOvr ? null : { ...MAPPING };
		},
		read: (chart) => chart.clrMapOvr ?? null,
	},
	{
		name: 'colour style palette',
		applies: (chart) => chart.colorPalette !== undefined && chart.colorStylePartPath !== undefined,
		edit: (chart) => {
			chart.colorPalette = ['#C00000', '#0070C0', '#00B050'];
			chart.colorMethod = 'cycle';
		},
		read: (chart) => chart.colorPalette,
	},
	{
		name: 'user shapes removed',
		applies: (chart) => (chart.userShapes?.length ?? 0) > 0,
		edit: (chart) => {
			chart.userShapes = [];
		},
		read: (chart) => chart.userShapes?.length ?? 0,
	},
];

/**
 * Both paths keep every edit; the colour style part is the same sibling part on both. The XML
 * differs only in what the object-tree writers do to the rest of an edited part (theme title font
 * pinning, JavaScript number caches, line series `c:idx`, dropped switched-off pie `c:dLbls`), as
 * recorded in `chart-group-flags-migration.test.ts`; the edited elements agree.
 */
const EXPECTED = {
	'external data auto-update': { parts: 35, legacyKept: 35, xml: 10 },
	'colour map override': { parts: 75, legacyKept: 75, xml: 30 },
	'colour style palette': { parts: 31, legacyKept: 31, xml: 8 },
	'user shapes removed': { parts: 1, legacyKept: 1, xml: 1 },
};

export function describePartSettings(...names: (keyof typeof EXPECTED)[]): Promise<void> {
	return describeFamilyMigrationEdits(EDITS, EXPECTED, names);
}
