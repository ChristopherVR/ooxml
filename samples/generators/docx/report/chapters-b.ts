/** Section 4, the results: the portrait part, then the monthly record on a landscape page. */
import type { Block } from '../../../../src/core/docx/index';
import { rich, t } from '../kit/builder';
import type { CellInput, CellSpec } from '../kit/tables';
import { energyChart, powerCurveChart } from './charts';
import type { Report } from './context';
import {
	LOSSES,
	MONTHS,
	n,
	PALETTE,
	POWER_CURVE,
	QUARTERS,
	REPORT,
	REVIEWERS,
	STEADY,
	YEAR,
} from './data';
import { TEXT_WIDTH, WIDE_WIDTH } from './styles';

const HEAD_FILL = `#${PALETTE.deep}`;
/** A header cell of a two-row table header: the table style only formats the first row. */
const head = (content: string, extra: Partial<CellSpec> = {}): CellInput => ({
	content: t(content, { bold: true, color: `#${PALETTE.white}` }),
	fill: HEAD_FILL,
	vAlign: 'bottom',
	...extra,
});
const percent = (value: number) => `${value.toFixed(1)}%`;

export function results(report: Report): Block[] {
	const { doc } = report;
	const { leona, hugo } = REVIEWERS;
	const commissioning = {
		measured: MONTHS[0]!.measured + MONTHS[1]!.measured,
		modelled: MONTHS[0]!.modelled + MONTHS[1]!.modelled,
	};
	const belowRated = POWER_CURVE.filter((bin) => bin.modelled > 0 && bin.modelled < REPORT.ratedKw);
	const shares = belowRated.map((bin) => Math.round((100 * bin.measured) / bin.modelled));
	const reached = POWER_CURVE.find((bin) => bin.measured >= 0.98 * REPORT.ratedKw);
	return [
		report.h(1, 'results', 'Results'),
		report.h(2, 'yield', 'Energy yield'),
		report.p(
			`The array delivered **${n(YEAR.measured)} MWh** to the grid in its first year, ${percent(YEAR.ratio)} of the ${n(YEAR.modelled)} MWh the model predicted{note}, for a capacity factor of ${percent(YEAR.capacityFactor)}. {figure} shows the monthly comparison and {table} the quarterly totals.`,
			{
				figure: report.fig('energy'),
				table: report.tab('quarterly'),
				note: doc.footnote(
					rich('Net of the 7.5% loss allowance described in {section}.', {
						section: report.sec('model'),
					}),
				),
			},
		),
		report.p(
			`Two periods account for almost all of the shortfall. During commissioning in November and December the turbines were energised one at a time, and the array produced ${n(commissioning.measured)} MWh against ${n(commissioning.modelled)} MWh predicted. In May, turbine T3 was recovered to replace a gearbox seal and the array ran on three machines for 26 days. Outside those two periods, in the eight months when availability was above 94%, the array delivered ${percent(STEADY.ratio)} of the modelled energy.`,
		),
		doc.p('Figure', [energyChart()]),
		report.figureCaption('energy', 'Modelled and measured net energy by month'),
		report.tableCaption('quarterly', 'Net energy and availability by quarter'),
		report.dataTable({
			grid: [1100, 1900, 1500, 1500, 1600, TEXT_WIDTH - 7600],
			headerRows: 2,
			look: {
				firstRow: true,
				lastRow: true,
				firstColumn: true,
				lastColumn: false,
				noHBand: false,
				noVBand: true,
			},
			columnAlign: [undefined, undefined, 'right', 'right', 'right', 'right'],
			rows: [
				[
					head('Quarter', { merge: 'restart' }),
					head('Months', { merge: 'restart' }),
					head('Net energy (MWh)', { span: 2, align: 'center' }),
					head('Measured over modelled', { merge: 'restart', align: 'right' }),
					head('Availability', { merge: 'restart', align: 'right' }),
				],
				[
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
					head('Modelled', { align: 'right' }),
					head('Measured', { align: 'right' }),
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
				],
				...QUARTERS.map((quarter) => [
					quarter.label,
					quarter.span,
					n(quarter.modelled),
					n(quarter.measured),
					percent(quarter.ratio),
					percent(quarter.availability),
				]),
				[
					'Year',
					'Nov to Oct',
					n(YEAR.modelled),
					n(YEAR.measured),
					percent(YEAR.ratio),
					percent(YEAR.availability),
				],
			],
		}),
		doc.p(
			'KeyFinding',
			rich(
				`**Key finding.** In steady operation the array delivers within 5% of the modelled energy. The year-one shortfall of ${n(YEAR.modelled - YEAR.measured)} MWh is an availability problem, not a resource problem.`,
			),
		),
		report.h(2, 'availability', 'Availability'),
		report.p(
			`Availability, the share of turbine-hours in which a machine was able to generate, was ${percent(YEAR.availability)} over the year and ${percent(STEADY.availability)} in the eight steady months. The ${n(YEAR.outage)} turbine-hours lost fall into three groups:`,
			{},
			{ keepNext: true },
		),
		...report.bullets([
			`**Commissioning**, November and December: ${n(LOSSES.commissioning)} turbine-hours, ${LOSSES.commissioningShare}% of the total.`,
			rich(
				`**The recovery of T3**, {comment}: ${n(LOSSES.seal)} turbine-hours, ${LOSSES.sealShare}%.`,
				{
					comment: doc.comment('May', [
						{
							...hugo,
							date: '2026-01-09T08:40:00Z',
							text: 'Add the vessel mobilisation time to this breakdown before issue; it is most of the 26 days.',
						},
					]),
				},
			),
			`**Everything else**, in ten months: ${n(LOSSES.other)} turbine-hours, ${LOSSES.otherShare}%, made up of a storm stand-down, planned inspections and converter trips.`,
		]),
		report.p(
			'Access is the constraint behind all three. A turbine can be reached only in the slack-water window, so even a short repair costs at least one tidal cycle, and the recovery of T3 waited nine days for a vessel{note}.{added}',
			{
				note: doc.endnote(
					'The recovery vessel, a crane barge shared with the Port Averly harbour works, was committed elsewhere from 3 to 11 May 2025. A call-off agreement signed in August 2025 guarantees mobilisation within 72 hours.',
				),
				added: doc.inserted(
					' A spare seal kit and a lifting frame are now held at the quay.',
					leona,
					'2026-01-08T10:55:00Z',
				),
			},
		),
		report.h(2, 'curve', 'Power curve'),
		report.p(
			`{figure} compares the binned electrical power of the four turbines with {equation}. Below rated speed the machines deliver between ${Math.min(...shares)}% and ${Math.max(...shares)}% of the modelled power; the gap is consistent with drivetrain and converter losses, which the model applies to the array as a whole rather than to the curve. Rated power is reached at ${reached?.speed.toFixed(1)} m/s rather than the 2.28 m/s predicted, and is held to within 1.2% above it.`,
			{ figure: report.fig('curve'), equation: report.eq('power') },
		),
		report.p(
			'The curve was built by the method of bins. Ten-minute means of the flow speed at hub height, taken from the upstream profiler, were sorted into bins 0.2 m/s wide, and the electrical power of each turbine was averaged within each bin. Only records with all four machines available and the flow within 10° of the rotor axis were used, which left about 21,500 records, or 41% of the year.',
		),
		report.p(
			'Blade-root strain on T1 and T3 stayed within 62% of the design envelope. The largest loads came in the first minutes after slack water, when the rotor yaws through the turn of the tide.',
		),
		doc.p('Figure', [powerCurveChart()]),
		report.figureCaption('curve', 'Power curve of one turbine: model and measurement'),
		report.h(2, 'environment', 'Environmental monitoring'),
		report.p(
			'The consent sets six indicators, each with a threshold that triggers a review. All six stayed within their thresholds ({table}). The most closely watched is underwater noise: the array raised the broadband level at 100 m by 6 dB over the baseline, well under the level at which harbour porpoises are expected to avoid the channel {cite}. Porpoise detections fell by 7%, which is inside the year-to-year variation of the three baseline years.',
			{
				table: report.tab('environment'),
				cite: report.cite('Cas24', '(Castellanos & Dahlgren, 2024)'),
			},
		),
		report.tableCaption('environment', 'Environmental indicators against their consent thresholds'),
		report.dataTable({
			grid: [3900, 1200, 1200, 1450, TEXT_WIDTH - 7750],
			columnAlign: [undefined, 'right', 'right', 'right', undefined],
			rows: [
				[
					'Indicator',
					{ content: 'Baseline', align: 'right' },
					{ content: 'Year one', align: 'right' },
					{ content: 'Threshold', align: 'right' },
					'Status',
				],
				['Broadband noise at 100 m (dB re 1 µPa)', '118', '124', '135', 'Within'],
				[
					'Porpoise detections (click-positive minutes per day)',
					'41',
					'38',
					'at least 30',
					'Within',
				],
				['Scour depth at the foundations (m)', '0.00', '0.21', '0.50', 'Within'],
				['Benthic richness (taxa per grab sample)', '23', '22', 'at least 18', 'Within'],
				['Median turbidity (NTU)', '4.1', '4.3', '6.0', 'Within'],
				['Recorded collisions', 'none', 'none', 'none', 'Within'],
			],
		}),
	];
}

/** Section 4.5 and its wide table: the content of the landscape section. */
export function monthlyRecord(report: Report): Block[] {
	const { doc } = report;
	const fixed = [1100, 700, 1100, 1150, 850, 1250, 1050, 850, 850, 1250];
	const notes = WIDE_WIDTH - fixed.reduce((sum, width) => sum + width, 0);
	const right = { align: 'right' as const };
	return [
		report.h(2, 'monthly', 'Monthly record'),
		report.p(
			'{table} gives the complete record on which this section is based. Energy is net at the fiscal meter; outage is counted in turbine-hours, so a full day with one machine out of service is 24.',
			{ table: report.tab('monthly') },
		),
		report.tableCaption(
			'monthly',
			'Monthly performance of the array, November 2024 to October 2025',
		),
		report.dataTable({
			grid: [...fixed, notes],
			headerRows: 2,
			look: {
				firstRow: true,
				lastRow: true,
				firstColumn: true,
				lastColumn: false,
				noHBand: false,
				noVBand: true,
			},
			columnAlign: [
				undefined,
				'right',
				'right',
				'right',
				'right',
				'right',
				'right',
				'right',
				'right',
				'right',
				undefined,
			],
			rows: [
				[
					head('Month', { merge: 'restart' }),
					head('Days', { merge: 'restart', ...right }),
					head('Net energy (MWh)', { span: 2, align: 'center' }),
					head('Ratio', { merge: 'restart', ...right }),
					head('Availability', { merge: 'restart', ...right }),
					head('Capacity factor', { merge: 'restart', ...right }),
					head('Flow (m/s)', { span: 2, align: 'center' }),
					head('Outage (turbine-h)', { merge: 'restart', ...right }),
					head('Notes', { merge: 'restart' }),
				],
				[
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
					head('Modelled', right),
					head('Measured', right),
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
					head('Mean', right),
					head('Peak', right),
					head('', { merge: 'continue' }),
					head('', { merge: 'continue' }),
				],
				...MONTHS.map((month) => [
					`${month.label} ${month.long.slice(-4)}`,
					String(month.days),
					n(month.modelled),
					n(month.measured),
					percent(month.ratio),
					percent(month.availability),
					percent(month.capacityFactor),
					month.meanFlow.toFixed(2),
					month.peakFlow.toFixed(2),
					n(month.outage),
					month.note,
				]),
				[
					'Year',
					String(YEAR.days),
					n(YEAR.modelled),
					n(YEAR.measured),
					percent(YEAR.ratio),
					percent(YEAR.availability),
					percent(YEAR.capacityFactor),
					YEAR.meanFlow.toFixed(2),
					YEAR.peakFlow.toFixed(2),
					n(YEAR.outage),
					'Day-weighted mean; highest peak',
				],
			],
		}),
		doc.p(
			'TableNote',
			'Ratio is measured over modelled net energy. Availability is the share of turbine-hours able to generate. Capacity factor is net energy over the energy at rated power for the whole month.',
		),
	];
}
