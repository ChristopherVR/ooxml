/** Sections 5 and 6, the references and the notes. */
import type { Block, ContentControl } from '../../../../src/core/docx/index';
import { rich, t } from '../kit/builder';
import { equation } from '../kit/drawing';
import { fieldClose, fieldOpen } from '../kit/fields';
import { tw } from '../kit/tables';
import { FONTS } from '../../lib/fiction';
import type { Report } from './context';
import { COST, n, REVIEWERS, STEADY, YEAR } from './data';
import { NUM, TEXT_WIDTH } from './styles';
import { SOURCES } from './sources';

export function discussion(report: Report): Block[] {
	const { doc } = report;
	const recalibrated = Math.round((YEAR.modelled * (1 - 0.019)) / 10) * 10;
	return [
		report.h(1, 'discussion', 'Discussion'),
		report.h(2, 'lessons', 'What the first year shows'),
		report.p(
			`The resource was never in doubt, and the first year confirms it: with the machines available, the array tracks the harmonic model to within five percent, month after month. This is the central result of the pilot. It means that the energy of a tidal array can be forecast years ahead with a confidence that wind and solar projects cannot offer, and that the remaining risk sits where engineers know how to address it, in keeping machines running.`,
		),
		doc.p('Quote', [
			t(
				'“We can tell you to the minute when the tide will turn in 2045. What we could not tell you, a year ago, was how long a gearbox seal would last.”',
			),
			t('\n'),
			t('Idris Okonjo, operations lead', { italic: false, fontSize: 9, fontFamily: FONTS.body }),
		]),
		report.p(
			`The same record shows where the model is optimistic. It assumes rated power from 2.28 m/s; the turbines reach it at 2.4 m/s. Recalibrating the curve lowers the predicted annual energy by 1.9%, to about ${n(recalibrated)} MWh. With the recalibrated curve and the 97.6% availability the operations plan now targets, the expected steady-state yield is ${n(COST.steadyEnergy)} MWh a year, against the ${STEADY.availability.toFixed(1)}% availability already achieved in the eight steady months.`,
		),
		report.h(2, 'cost', 'Levelised cost of energy'),
		report.p(
			'The levelised cost of energy (LCOE) spreads the lifetime cost of the array over the energy it is expected to produce, both discounted to present value {cite}:',
			{ cite: report.cite('Fal25', '(Falkenrath, Diallo, & Whitlow, 2025)') },
			{ keepNext: true },
		),
		report.equation(
			'lcoe',
			String.raw`\mathrm{LCOE} = \frac{\sum_{t=0}^{n} {\frac{I_t + M_t}{\left(1+r\right)^t}}}{\sum_{t=1}^{n} {\frac{E_t}{\left(1+r\right)^t}}}`,
		),
		report.p(
			'where {I} is the investment in year {t}, {M} the operating cost, {E} the net energy, {r} the discount rate and {n} the design life in years.',
			{
				I: equation('I_t'),
				t: equation('t'),
				M: equation('M_t'),
				E: equation('E_t'),
				r: equation('r'),
				n: equation('n'),
			},
		),
		report.p(
			`{table} lists the inputs. With the year-one cost record and the steady-state yield above, the LCOE of the pilot is {comment}{note}. That is roughly three times the cost of new offshore wind in the region, as expected of a four-machine pilot that carries the full cost of a recovery vessel, a cable and a consent. The sensitivity is instructive: every percentage point of availability is worth about €3.30/MWh, and sharing the vessel and the cable across a twelve-turbine array would bring the figure under €200/MWh on the institute’s estimate{end}.`,
			{
				table: report.tab('cost'),
				comment: doc.comment(t(`€${COST.lcoe}/MWh`, { bold: true }), [
					{
						...REVIEWERS.leona,
						date: '2026-01-08T11:20:00Z',
						text: 'Please confirm the 6.5% real discount rate against the funder’s latest guidance before issue.',
					},
				]),
				note: doc.footnote(
					'All costs are in 2025 euros and exclude grant funding. The capital cost includes the consent and the pre-construction surveys.',
				),
				end: doc.endnote(
					'Estimate prepared for the Northwind Coast Marine Energy Programme, assuming twelve MB-500 turbines on the existing cable corridor and one shared recovery vessel. It has not been independently audited.',
				),
			},
		),
		report.tableCaption('cost', 'Inputs to the levelised cost of energy and the result'),
		report.dataTable({
			grid: [2900, 2300, TEXT_WIDTH - 5200],
			look: {
				firstRow: true,
				lastRow: true,
				firstColumn: true,
				lastColumn: false,
				noHBand: false,
				noVBand: true,
			},
			rows: [
				['Input', 'Value', 'Basis'],
				[
					'Capital cost',
					`€${COST.capex.toFixed(1)} million`,
					'Outturn, including foundations and cable',
				],
				[
					'Operating cost',
					`€${COST.opex.toFixed(2)} million a year`,
					'Year-one actuals, less commissioning',
				],
				[
					'Discount rate',
					`${(COST.rate * 100).toFixed(1)}% real`,
					'Guidance of the funding programme',
				],
				['Design life', `${COST.life} years`, 'Type certificate'],
				['Annual net energy', `${n(COST.steadyEnergy)} MWh`, 'Steady-state yield, this section'],
				[
					'Present value of costs',
					`€${COST.presentCost.toFixed(1)} million`,
					'Numerator of the LCOE expression',
				],
				[
					'Present value of energy',
					`${n(COST.presentEnergy)} MWh`,
					'Denominator of the LCOE expression',
				],
				['Levelised cost of energy', `€${COST.lcoe}/MWh`, 'Costs in 2025 euros, before grants'],
			],
		}),
	];
}

export function conclusions(report: Report): Block[] {
	const { doc } = report;
	const numbered = (markup: string, first = false) =>
		doc.p('ListNumber', rich(markup), first ? { numbering: { numId: NUM.steps, level: 0 } } : {});
	return [
		report.h(1, 'conclusions', 'Conclusions and recommendations'),
		report.h(2, 'findings', 'Conclusions'),
		numbered(
			`The array delivered ${n(YEAR.measured)} MWh in its first year, a capacity factor of ${YEAR.capacityFactor.toFixed(1)}%.`,
			true,
		),
		numbered(
			`In the eight months of steady operation, measured energy was within 5% of the pre-construction model. The harmonic resource model is validated for the Narrows.`,
		),
		numbered(
			`The shortfall against the model, ${n(YEAR.modelled - YEAR.measured)} MWh, was caused by commissioning and by one drivetrain recovery, not by the resource.`,
		),
		numbered('All six environmental indicators stayed within their consent thresholds.'),
		numbered(
			`The levelised cost of energy of the pilot is €${COST.lcoe}/MWh. Availability and shared infrastructure are the two largest levers on it.`,
		),
		report.h(2, 'recommendations', 'Recommendations'),
		...report.bullets([
			'Recalibrate the power curve in the resource model to the measured rated speed of 2.4 m/s before the year-two forecast is issued.',
			'Hold critical drivetrain spares at the quay and keep the vessel call-off agreement in force.',
			'Extend the hydrophone record through a second porpoise calving season before any expansion of the array is proposed.',
			'Publish the ten-minute flow and power record as an open dataset, so that other channels on the Northwind Coast can be assessed against it.',
		]),
		report.p('The dataset and this report are available from {link}.', {
			link: t('saltmarsh.example/greywharf', {
				style: 'Hyperlink',
				link: { href: 'https://saltmarsh.example/greywharf', tooltip: 'Greywharf pilot data' },
			}),
		}),
	];
}

/** The references: a BIBLIOGRAPHY field in a bibliography content control, entries in place. */
export function references(report: Report): Block[] {
	const { doc } = report;
	const control: ContentControl = { id: '880001', level: 'block', kind: 'bibliography' };
	const entries = SOURCES.map((source, index) =>
		doc.p(
			'Bibliography',
			[...(index === 0 ? fieldOpen('BIBLIOGRAPHY') : []), ...rich(source.entry)],
			{ contentControl: control },
		),
	);
	// As Word writes it, the field ends in a paragraph of its own after the last entry.
	const end = doc.p('Bibliography', [fieldClose()], {
		contentControl: control,
		spacingAfterTwips: tw(0),
		lineSpacingTwips: tw(40),
		lineSpacingRule: 'exact',
	});
	return [report.plainHeading('references', 'References'), ...entries, end];
}

/** The last heading of the document: Word lays the endnotes out after it. */
export function notes(report: Report): Block[] {
	return [report.plainHeading('notes', 'Notes')];
}
