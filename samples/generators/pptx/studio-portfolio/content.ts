/**
 * The portfolio's facts in one place, so the case slides, the results table and the speaker
 * notes all quote the same figures. Everything here is invented.
 */
import { PEOPLE } from '../../lib/fiction';

export interface CaseStudy {
	number: string;
	client: string;
	discipline: string;
	year: number;
	/** One line under the thumbnail on the index slide. */
	teaser: string;
	/** The caption under the large image on the opener. */
	caption: string;
	/** The detail slide's headline and brief. */
	headline: string;
	brief: string;
	made: [string, string, string];
	figures: [value: string, caption: string][];
	/** The results table row. */
	scope: string;
	weeks: number;
	result: string;
	notes: { opener: string; detail: string };
}

export const CASES: CaseStudy[] = [
	{
		number: '01',
		client: 'Brindlemoor Bakehouse',
		discipline: 'Identity and packaging',
		year: 2023,
		teaser: 'One wheat mark for a bakery group growing from nine shops to fourteen.',
		caption:
			'The wheat ear becomes the mark, the round stamp and the stripe on every awning: three parts a baker can apply without calling us.',
		headline: 'One mark, baked into everything.',
		brief:
			'Brindlemoor was opening five shops in a year and looked different in each one. We drew a single wheat mark and a kit of parts the shop teams can apply themselves.',
		made: [
			'A wheat mark, a round stamp and a three-colour palette',
			'62 packaging formats, from loaf bags to celebration cake boxes',
			'Shopfront, awning and menu-board guidelines for 14 bakeries',
		],
		figures: [
			['+18%', 'repeat purchases within 30 days, year on year'],
			['54%', 'unprompted brand recall in its home region, up from 31%'],
			['16', 'weeks from brief to the first rebranded shop'],
		],
		scope: '14 bakeries, 62 packaging formats',
		weeks: 16,
		result: 'Repeat purchases up 18%',
		notes: {
			opener:
				'Case one is Brindlemoor Bakehouse, a regional bakery group that was growing from nine shops to fourteen. This board is the whole identity on one sheet: the wheat mark, the stamp, the awning stripe and the loaf pictogram. The image grew out of the thumbnail on the previous slide; that is a Morph transition.',
			detail:
				'The brief was consistency without a brand police. We delivered one mark, 62 packaging formats and guidelines for all 14 bakeries in 16 weeks. In the first year repeat purchases within 30 days rose 18 percent and unprompted recall went from 31 to 54 percent.',
		},
	},
	{
		number: '02',
		client: 'Larkspur Heights Community Garden',
		discipline: 'Wayfinding',
		year: 2024,
		teaser: 'A sign family that volunteers can repaint, extend and repair themselves.',
		caption:
			'Three totem sizes share one grid: a coloured head with a pictogram, an arrow, then two lines of text. Nothing else is allowed on a sign.',
		headline: 'Find your plot without asking.',
		brief:
			'The garden had grown to 3.2 hectares and 180 plots, and new members kept getting lost. We designed a sign family the volunteers can paint, extend and repair on their own.',
		made: [
			'48 signs in four sizes, from gate totems to plot markers',
			'22 pictograms drawn on one 24-unit grid',
			'A stencil kit and repainting schedule for the volunteer team',
		],
		figures: [
			['91%', 'of new members found their plot unaided, up from 46%'],
			['+37%', 'volunteer sign-ups in the first season after installation'],
			['11', 'weeks from site survey to the last sign in the ground'],
		],
		scope: '48 signs, 22 pictograms',
		weeks: 11,
		result: 'Plots found unaided: 91%, was 46%',
		notes: {
			opener:
				'Case two is the Larkspur Heights Community Garden, a neighbourhood association rather than a company. The picture shows the three totem sizes standing along the main path. Every sign follows the same order: colour, pictogram, arrow, two lines of text.',
			detail:
				'The garden covers 3.2 hectares and 180 plots. We made 48 signs and 22 pictograms in 11 weeks and left a stencil kit behind so the volunteers own the system. In the survey after the first season, 91 percent of new members found their plot unaided, up from 46 percent, and volunteer sign-ups rose 37 percent.',
		},
	},
	{
		number: '03',
		client: 'Kestrelline Outfitters',
		discipline: 'Editorial',
		year: 2025,
		teaser: 'A catalogue rebuilt as a field guide to 38 walks, with the kit for each.',
		caption:
			'The cover of the autumn edition: six ridgelines, a contour grid and one dotted route. The same drawing system runs through all 164 pages.',
		headline: 'A catalogue kept in the rucksack.',
		brief:
			'Kestrelline sent out a price list twice a year. We rebuilt it as a field guide: 38 mapped routes, each printed beside the equipment it needs.',
		made: [
			'A 164-page field guide in two seasonal editions',
			'38 route maps drawn on one contour system',
			'A type and colour system shared by print, web and swing tags',
		],
		figures: [
			['+24%', 'orders traced to the catalogue, edition on edition'],
			['9.1%', 'product return rate, down from 12.4%'],
			['20', 'weeks from the first route walk to press'],
		],
		scope: '164 pages, 38 route maps',
		weeks: 20,
		result: 'Catalogue orders up 24%',
		notes: {
			opener:
				'Case three is Kestrelline Outfitters, an outdoor equipment retailer. This is the cover of the autumn field guide, shown full bleed. Notice the dotted route on the nearest ridge: it is the first of the 38 walks in the book.',
			detail:
				'Kestrelline used to post a price list. The field guide has 164 pages and 38 route maps and took 20 weeks from the first route walk to press. Orders traced to the catalogue rose 24 percent, and returns fell from 12.4 to 9.1 percent because people bought the right kit for the walk.',
		},
	},
];

/** Weeks across the three projects (16 + 11 + 20); each lies within the phases' 11 to 20. */
export const TOTAL_WEEKS = CASES.reduce((sum, study) => sum + study.weeks, 0);

/** Projects shipped from 2023 to 2025, by discipline; they add up to `PROJECTS`. */
export const MIX: [discipline: string, projects: number][] = [
	['Identity', 15],
	['Editorial', 11],
	['Wayfinding', 9],
	['Packaging', 6],
];
export const PROJECTS = MIX.reduce((sum, [, count]) => sum + count, 0);

export const STUDIO = {
	people: 9,
	founded: 2016,
	city: 'Port Averly',
	email: 'hello@thimblewick.example',
	site: 'thimblewick.example',
	address: '12 Ropewalk Lane, Port Averly',
};

/** The five of the nine a client talks to; names from the shared cast. */
export const TEAM: [name: string, role: string][] = [
	[PEOPLE[11], 'Founder and creative director'],
	[PEOPLE[3], 'Design director, identity'],
	[PEOPLE[6], 'Wayfinding lead'],
	[PEOPLE[7], 'Editorial designer'],
	[PEOPLE[15], 'Studio producer'],
];

export const PHASES: [name: string, share: string, text: string][] = [
	[
		'Listen',
		'1 to 2 weeks',
		'On site before we draw: counting signs, shadowing staff, reading the returns.',
	],
	[
		'Frame',
		'2 to 3 weeks',
		'One page that says what the work must do, agreed before any design is shown.',
	],
	['Make', '6 to 13 weeks', 'Prototypes at full size, tested with the people who will use them.'],
	[
		'Hand over',
		'2 weeks',
		'Guidelines, source files and a training day, so you never need to call us.',
	],
];
