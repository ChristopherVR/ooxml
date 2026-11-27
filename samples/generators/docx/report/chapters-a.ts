/** Sections 1 to 3: introduction, the site and installation, and the methods. */
import type { Block } from '../../../../src/core/docx/index';
import { rich, t } from '../kit/builder';
import { equation, pxFromCm } from '../kit/drawing';
import { flowMap, sitePlan } from './art';
import type { Report } from './context';
import { n, PALETTE, RATED_SPEED, REPORT, REVIEWERS, ROTOR, SWEPT_AREA, YEAR } from './data';
import { TEXT_WIDTH } from './styles';

/** A label cell of a table that pairs labels with values. */
const LABEL = { bold: true, color: `#${PALETTE.deep}` } as const;

export function introduction(report: Report): Block[] {
	const { doc } = report;
	return [
		report.h(1, 'intro', 'Introduction'),
		report.h(2, 'background', 'Background'),
		report.p(
			'Greywharf sits at the mouth of the Averly estuary, where the tide is forced through a channel barely four hundred metres wide. Twice a day the Narrows carries more water than every river on the Northwind Coast combined, and it does so on a timetable that can be printed decades ahead. The Saltmarsh Institute for Coastal Engineering began measuring that flow in 2021 {cite}, and in November 2024 it energised the first grid-connected tidal array in the region: four 500 kW turbines on gravity foundations, exporting to the Greywharf quay substation.',
			{ cite: report.cite('Ren23', '(Renner & Okonjo, 2023)') },
		),
		report.p(
			'The pilot was built to answer three practical questions. How closely does a harmonic resource model predict what a real array delivers? What does it take to keep machines available in a channel that can be worked for less than an hour around slack water? And does the array disturb the marine mammals, the seabed or the water quality of a designated estuary?{note}',
			{
				note: doc.footnote(
					rich(
						'The Averly estuary was designated a marine conservation area by the Northwind Coast Environment Board in 2019. The consent for the pilot sets the review thresholds summarised in {table}.',
						{ table: report.tab('environment') },
					),
				),
			},
		),
		report.h(2, 'scope', 'Scope and objectives'),
		report.p(
			`This report covers the first full year of operation, ${REPORT.period}. It sets out:`,
			{},
			{ keepNext: true },
		),
		...report.bullets([
			rich(
				'the net energy delivered to the grid, compared month by month with the pre-construction model ({s});',
				{ s: report.sec('yield') },
			),
			rich('the availability of the four turbines and the causes of lost production ({s});', {
				s: report.sec('availability'),
			}),
			rich('the measured power curve of the machines ({s});', { s: report.sec('curve') }),
			rich('the results of the environmental monitoring programme ({s}); and', {
				s: report.sec('environment'),
			}),
			rich('an updated estimate of the levelised cost of energy ({s}).', { s: report.sec('cost') }),
		]),
		report.h(2, 'reading', 'How to read this report'),
		report.p(
			'{site} describes the site and the installation, and {methods} the methods. The results follow in {results}, with the complete monthly record in {table} on page {page}. Readers interested mainly in the findings may go straight to {conclusions}. Energy is given in megawatt-hours (MWh) net of array losses, flow speeds are depth-averaged unless stated otherwise, and all times are local standard time.',
			{
				site: report.sec('site'),
				methods: report.sec('methods'),
				results: report.sec('results'),
				table: report.tab('monthly'),
				page: report.page('tab', 'monthly'),
				conclusions: report.sec('conclusions'),
			},
		),
	];
}

export function site(report: Report): Block[] {
	const { doc } = report;
	return [
		report.h(1, 'site', 'Site and installation'),
		report.h(2, 'narrows', 'The Greywharf Narrows'),
		report.p(
			`The channel runs east to west for about 2.4 km between the Greywharf quay and Thistledown Point. At its throat it is 380 m wide and 32 m deep at mean sea level, with a bed of swept bedrock and coarse gravel ({figure}). The tide is semi-diurnal, with a mean spring range of 4.6 m. Over the reporting year the depth-averaged flow at the array averaged ${YEAR.meanFlow.toFixed(2)} m/s and peaked at ${YEAR.peakFlow.toFixed(2)} m/s during the equinoctial springs of March 2025.`,
			{ figure: report.fig('site') },
		),
		doc.p('Figure', [
			doc.picture(sitePlan(), {
				name: 'site-plan',
				widthPx: pxFromCm(15),
				heightPx: pxFromCm(8.4),
				alt: 'Plan of a narrowing tidal channel between two shores, with four turbines in the throat, a cable to the north shore and a current profiler on each side.',
				title: 'Site plan',
			}),
		]),
		report.figureCaption('site', 'Plan of the Greywharf Narrows and the pilot array'),
		doc.p(
			'TableNote',
			'Amber: turbines T1 to T4. Diamonds: current profilers. Red: export cable. Chevrons: flood tide. Scale bar: 500 m.',
		),
		report.h(2, 'array', 'Turbine array'),
		report.p(
			'The array consists of four MB-500 turbines supplied by Brackwater Marine Works{note}, arranged in two staggered rows across the throat of the channel. Each machine has a three-bladed rotor that yaws to face the flood and the ebb, a two-stage gearbox and a permanent-magnet generator. {table} lists the principal particulars.',
			{
				table: report.tab('spec'),
				note: doc.endnote(
					'The MB-500 was type-tested at the Marrow Bay tow tank in 2023. The certificate covers flow speeds up to 4.0 m/s and a design life of 20 years.',
				),
			},
		),
		report.tableCaption('spec', 'Principal particulars of the array'),
		// Two label and value pairs to a row keep the table short enough to sit under the figure.
		report.dataTable({
			grid: [1500, 3050, 1950, TEXT_WIDTH - 6500],
			rows: [
				['Item', 'Particulars', 'Item', 'Particulars'],
				...(
					[
						[
							'Turbines',
							`${REPORT.turbines} × MB-500, horizontal axis`,
							'Hub height',
							'14 m above the bed',
						],
						[
							'Rated power',
							`${REPORT.ratedKw} kW each, ${YEAR.ratedMw.toFixed(1)} MW in all`,
							'Cut-in, rated speed',
							`${ROTOR.cutIn.toFixed(1)} m/s, ${RATED_SPEED.toFixed(2)} m/s`,
						],
						[
							'Rotor',
							`${ROTOR.diameter} m, three blades, ${SWEPT_AREA} m²`,
							'Foundation',
							'Gravity base of 310 t',
						],
						[
							'Export cable',
							'11 kV, 620 m, armoured',
							'Grid connection',
							'Quay substation, 2.5 MVA',
						],
					] as const
				).map(([item, value, other, otherValue]) => [
					t(item, LABEL),
					value,
					t(other, LABEL),
					otherValue,
				]),
			],
		}),
		report.h(3, 'rotor', 'Rotor and drivetrain'),
		report.p(
			`The blades are a glass and carbon fibre laminate with a pitch range of 95°, which lets the controller hold rated power above ${RATED_SPEED.toFixed(2)} m/s and feather the rotor before a recovery. The nacelle is a sealed, oil-filled unit: the gearbox and generator share one lubrication circuit, and a single wet-mate connector carries power and data to the foundation.`,
		),
		report.h(3, 'foundations', 'Foundations and cable'),
		report.p(
			'Each gravity base was lowered in a single slack-water window{note} and levelled on three hydraulic feet. No drilling or piling was needed, which shortened the marine campaign to eleven days and kept installation noise below the threshold set for the porpoise calving season. The export cable is bolted to the bedrock at two-metre intervals for its first 180 m and runs in a rock-cut trench to the quay.',
			{
				note: doc.footnote(
					'Slack water in the Narrows, defined as flow below 0.5 m/s, lasts between 38 and 64 minutes depending on the phase of the spring-neap cycle.',
				),
			},
		),
		report.h(2, 'grid', 'Grid connection'),
		report.p(
			'The array exports through a single 11 kV cable to the quay substation, where a 2.5 MVA transformer and a static compensator meet the reactive power and fault ride-through requirements of the regional connection code {cite}. Net energy in this report is the reading of the fiscal meter on the grid side of that transformer.',
			{ cite: report.cite('Nor24', '(Northwind Coast Grid Authority, 2024)') },
		),
	];
}

export function methods(report: Report): Block[] {
	const { doc } = report;
	const { leona, hugo } = REVIEWERS;
	return [
		report.h(1, 'methods', 'Methods'),
		report.h(2, 'model', 'Resource model'),
		report.p(
			'The pre-construction energy estimate rests on a harmonic model of the current. The depth-averaged flow speed is written as a sum of {N} tidal constituents,',
			{ N: equation('N') },
			{ keepNext: true },
		),
		report.equation(
			'harmonic',
			String.raw`v(t) = \sum_{i=1}^{N} {a_i \cos\left(\omega_i t + \varphi_i\right)}`,
		),
		report.p(
			'where {a}, {w} and {phi} are the amplitude, angular frequency and phase of constituent {i}. Fifteen constituents were fitted to 14 months of profiler data {cite}; the two largest, the principal lunar and principal solar semi-diurnal constituents{note}, account for 89% of the variance. {figure} shows the modelled speed at peak spring flood: the flow more than doubles through the throat, and the array sits in the band above 3 m/s.',
			{
				a: equation('a_i'),
				w: equation(String.raw`\omega_i`),
				phi: equation(String.raw`\varphi_i`),
				i: equation('i'),
				cite: report.cite('Vas22', '(Vasanth, 2022)'),
				figure: report.fig('flow'),
				note: doc.footnote(
					'The principal lunar constituent has a period of 12.42 hours and the principal solar constituent one of 12.00 hours. Their beat produces the spring-neap cycle of 14.77 days.',
				),
			},
		),
		doc.p('Figure', [
			doc.picture(flowMap(), {
				name: 'flow-map',
				widthPx: pxFromCm(15),
				heightPx: pxFromCm(8.4),
				alt: 'Map of modelled flow speed in the channel: slow near the shores and at the ends, fastest in the narrow throat where the turbines stand.',
				title: 'Flow map',
			}),
		]),
		report.figureCaption('flow', 'Modelled depth-averaged flow speed at peak spring flood'),
		doc.p(
			'TableNote',
			'The legend at the lower left runs from 0 to 3.5 m/s in bands of 0.5 m/s; the amber band is flow above 3 m/s. Turbine positions are marked as in Figure 1.',
		),
		report.p(
			'The hydrodynamic power a rotor of swept area {A} can extract from a flow of speed {v} is',
			{ A: equation('A'), v: equation('v') },
			{ keepNext: true },
		),
		report.equation('power', String.raw`P = \frac{1}{2}\,\rho\,A\,C_p\,v^3`),
		report.p(
			`with {rho} the density of sea water (${n(ROTOR.density)} kg/m³) and {cp} the power coefficient, {comment} {cite}. Because power rises with the cube of speed, a 5% error in {v} becomes a 16% error in {P}, which is why the flow measurements of {section} matter more than any other input.`,
			{
				rho: equation(String.raw`\rho`),
				cp: equation('C_p'),
				v: equation('v'),
				P: equation('P'),
				cite: report.cite('Mor21', '(Morikawa, 2021)'),
				section: report.sec('instruments'),
				comment: doc.comment(`taken as ${ROTOR.powerCoefficient} from the supplier’s tank tests`, [
					{
						...leona,
						date: '2026-01-08T10:42:00Z',
						text: 'Is the tank-test coefficient corrected for blockage? The Narrows blocks about 4% of the cross-section at the array.',
					},
					{
						...hugo,
						date: '2026-01-09T08:15:00Z',
						text: 'Yes. Tobias applied the standard blockage correction; the uncorrected value was 0.43. Worth a footnote in the next issue.',
					},
				]),
			},
		),
		report.p(
			`Predicted energy is the integral of {equation} over the year, limited to rated power and reduced by 7.5% for wake, electrical and metering losses. The year-one prediction made this way was ${n(YEAR.modelled)} MWh, a capacity factor of ${YEAR.modelledCapacityFactor.toFixed(1)}%, where`,
			{ equation: report.eq('power') },
			{ keepNext: true },
		),
		report.equation(
			'cf',
			String.raw`\mathrm{CF} = \frac{E_{\mathrm{net}}}{P_{\mathrm{rated}} \cdot T}`,
		),
		report.p('and {T} is the length of the period.', { T: equation('T') }),
		report.h(2, 'instruments', 'Instrumentation'),
		report.p(
			'Two seabed-mounted acoustic Doppler current profilers bracket the array, 310 m upstream and 290 m downstream on the flood. Each turbine reports electrical power, rotor speed, pitch and yaw, and turbines T1 and T3 carry blade-root strain gauges. A four-element hydrophone array on the north shore listens both for turbine noise and for the echolocation clicks of harbour porpoises. {table} summarises the instruments.',
			{ table: report.tab('instruments') },
		),
		report.tableCaption('instruments', 'Instruments and sampling'),
		report.dataTable({
			grid: [2350, 2700, 2300, TEXT_WIDTH - 7350],
			rows: [
				['Instrument', 'Measures', 'Location', 'Sampling'],
				[
					'Current profilers (2)',
					'Flow speed and direction in 1 m bins',
					'Seabed, up- and downstream',
					'1 Hz, 10 min means',
				],
				['Power transducers (4)', 'Active and reactive power', 'Each nacelle', '10 Hz'],
				['Fiscal meter', 'Net exported energy', 'Quay substation', '30 min'],
				['Blade strain gauges', 'Flapwise and edgewise bending', 'T1 and T3 blade roots', '50 Hz'],
				[
					'Hydrophone array',
					'Noise level and porpoise clicks',
					'North shore, 100 m from T1',
					'96 kHz',
				],
				['Seabed camera lander', 'Scour and benthic cover', 'Foundation of T2', 'Quarterly'],
			],
		}),
		report.h(2, 'quality', 'Data quality'),
		report.p(
			'Ten-minute records were rejected where a profiler returned fewer than 80% good pings or where its two beam pairs disagreed by more than 0.15 m/s. After screening, {old}{new} of the flow record and 99.4% of the power record remained. Gaps in flow were filled from the harmonic model; gaps in power were not filled and are {format}. The scatter of the accepted flow samples within a speed bin is described by the sample standard deviation,',
			{
				old: doc.deleted('roughly 97%', hugo, '2026-01-09T08:21:00Z'),
				new: doc.inserted('96.8%', hugo, '2026-01-09T08:21:00Z'),
				format: doc.reformatted(
					{ text: 'counted as lost production', bold: true },
					hugo,
					'2026-01-09T08:24:00Z',
				),
			},
			{ keepNext: true },
		),
		report.equation(
			'sigma',
			String.raw`\sigma_v = \sqrt{\frac{1}{N-1}\sum_{k=1}^{N}{\left(v_k - \bar{v}\right)^2}}`,
		),
		report.p('which averaged 0.09 m/s across the bins used for the power curve of {section}.', {
			section: report.sec('curve'),
		}),
	];
}
