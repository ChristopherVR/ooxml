/** The Greywharf pilot's plan: workstreams, their tasks, the milestones and what waits on what. */
import type { HueName } from '../lib/palette';

export interface Task {
	key: string;
	name: string;
	start: string;
	finish: string;
	owner: string;
	/** Thousands. */
	budget: number;
	status?: 'Under way' | 'Complete';
}
export interface Row {
	name: string;
	hue: HueName;
	tasks: readonly Task[];
}

export const ROWS: readonly Row[] = [
	{
		name: 'Consenting',
		hue: 'plum',
		tasks: [
			{
				key: 'licence',
				name: 'Seabed licence',
				start: '2026-01-05',
				finish: '2026-03-13',
				owner: 'Leona Castellanos',
				budget: 85,
				status: 'Under way',
			},
			{
				key: 'environment',
				name: 'Environmental monitoring plan',
				start: '2026-04-01',
				finish: '2026-09-11',
				owner: 'Sunniva Dahlgren',
				budget: 140,
			},
		],
	},
	{
		name: 'Design',
		hue: 'sea',
		tasks: [
			{
				key: 'concept',
				name: 'Concept and site survey',
				start: '2026-01-05',
				finish: '2026-04-30',
				owner: 'Felix Brandauer',
				budget: 260,
				status: 'Under way',
			},
			{
				key: 'design',
				name: 'Detailed design',
				start: '2026-05-04',
				finish: '2026-08-14',
				owner: 'Oskar Lindqvist',
				budget: 410,
			},
		],
	},
	{
		name: 'Foundations',
		hue: 'moss',
		tasks: [
			{
				key: 'frame',
				name: 'Foundation frame',
				start: '2026-10-05',
				finish: '2027-01-08',
				owner: 'Yuki Tanabe',
				budget: 620,
			},
			{
				key: 'seabed',
				name: 'Seabed works',
				start: '2027-01-11',
				finish: '2027-03-12',
				owner: 'Kenji Morikawa',
				budget: 380,
			},
		],
	},
	{
		name: 'Fabrication',
		hue: 'teal',
		tasks: [
			{
				key: 'turbine',
				name: 'Turbine and nacelle fabrication',
				start: '2026-09-01',
				finish: '2027-02-19',
				owner: 'Felix Brandauer',
				budget: 1850,
			},
		],
	},
	{
		name: 'Marine works',
		hue: 'amber',
		tasks: [
			{
				key: 'cable',
				name: 'Cable lay',
				start: '2027-02-01',
				finish: '2027-03-31',
				owner: 'Linnea Sorensen',
				budget: 540,
			},
			{
				key: 'install',
				name: 'Turbine installation',
				start: '2027-04-05',
				finish: '2027-06-18',
				owner: 'Kenji Morikawa',
				budget: 720,
			},
		],
	},
	{
		name: 'Commissioning',
		hue: 'coral',
		tasks: [
			{
				key: 'grid',
				name: 'Grid tests',
				start: '2027-05-17',
				finish: '2027-06-25',
				owner: 'Noor Haddad',
				budget: 160,
			},
			{
				key: 'commission',
				name: 'Commissioning',
				start: '2027-06-28',
				finish: '2027-08-20',
				owner: 'Idris Okonjo',
				budget: 230,
			},
		],
	},
	{
		name: 'Monitoring',
		hue: 'sea',
		tasks: [
			{
				key: 'monitor',
				name: 'Performance and wildlife monitoring',
				start: '2027-07-12',
				finish: '2027-12-31',
				owner: 'Amara Diallo',
				budget: 300,
			},
		],
	},
];

export const MILESTONES = [
	{ name: 'Seabed licence granted', date: '2026-03-16', owner: 'Leona Castellanos' },
	{ name: 'Design freeze', date: '2026-08-14', owner: 'Felix Brandauer' },
	{ name: 'Turbine delivered', date: '2027-02-19', owner: 'Felix Brandauer' },
	{ name: 'First power', date: '2027-07-09', owner: 'Idris Okonjo' },
	{ name: 'Pilot review', date: '2027-12-10', owner: 'Mara Ellingwood' },
] as const;

/** [from, to, how]: `over` leaves the right end and drops onto the next bar, `under` leaves below. */
export const DEPENDENCIES: readonly (readonly [string, string, 'over' | 'under'])[] = [
	['design', 'frame', 'over'],
	['design', 'turbine', 'over'],
	['seabed', 'install', 'over'],
	['turbine', 'install', 'over'],
	['cable', 'grid', 'under'],
	['install', 'commission', 'over'],
];
