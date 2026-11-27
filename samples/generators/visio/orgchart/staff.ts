/** The people of the Saltmarsh Institute for Coastal Engineering. All of them are invented. */
import type { HueName } from '../lib/palette';

export interface Post {
	name: string;
	title: string;
	location: string;
	/** `yyyy-mm-dd`. */
	started: string;
	fte?: number;
	status?: 'In post' | 'Vacant' | 'Seconded';
}

export interface Team {
	lead: Post;
	members: readonly Post[];
}

export interface Department {
	name: string;
	hue: HueName;
	head: Post;
	teams: readonly [Team, Team];
}

export const DIRECTOR: Post = {
	name: 'Mara Ellingwood',
	title: 'Director',
	location: 'Port Averly',
	started: '2019-09-02',
};
export const DIRECTORATE: { name: string; hue: HueName } = {
	name: "Director's office",
	hue: 'amber',
};

export const DEPARTMENTS: readonly Department[] = [
	{
		name: 'Coastal Modelling',
		hue: 'sea',
		head: {
			name: 'Tobias Renner',
			title: 'Head of Coastal Modelling',
			location: 'Port Averly',
			started: '2020-01-13',
		},
		teams: [
			{
				lead: {
					name: 'Hugo Falkenrath',
					title: 'Lead, Wave and Surge Models',
					location: 'Port Averly',
					started: '2020-06-01',
				},
				members: [
					{
						name: 'Anouk Verhaegen',
						title: 'Hydrodynamic modeller',
						location: 'Port Averly',
						started: '2021-03-15',
					},
					{
						name: 'Rafael Monteverde',
						title: 'Data assimilation scientist',
						location: 'Port Averly',
						started: '2023-10-02',
						fte: 0.8,
					},
				],
			},
			{
				lead: {
					name: 'Sunniva Dahlgren',
					title: 'Lead, Sediment Dynamics',
					location: 'Marrow Bay',
					started: '2021-02-08',
				},
				members: [
					{
						name: 'Esi Agyeman',
						title: 'Geomorphologist',
						location: 'Marrow Bay',
						started: '2022-09-05',
					},
					{
						name: 'Vacant',
						title: 'Coastal ecologist',
						location: 'Marrow Bay',
						started: '2026-02-02',
						status: 'Vacant',
					},
				],
			},
		],
	},
	{
		name: 'Field Operations',
		hue: 'moss',
		head: {
			name: 'Priya Vasanth',
			title: 'Head of Field Operations',
			location: 'Greywharf',
			started: '2019-11-04',
		},
		teams: [
			{
				lead: {
					name: 'Kenji Morikawa',
					title: 'Lead, Survey Vessels',
					location: 'Greywharf',
					started: '2020-04-20',
				},
				members: [
					{
						name: 'Linnea Sorensen',
						title: 'Vessel master',
						location: 'Greywharf',
						started: '2020-08-17',
					},
					{
						name: 'Marek Zielonka',
						title: 'Hydrographic surveyor',
						location: 'Greywharf',
						started: '2024-01-08',
					},
				],
			},
			{
				lead: {
					name: 'Amara Diallo',
					title: 'Lead, Instrumentation',
					location: 'Greywharf',
					started: '2021-07-12',
				},
				members: [
					{
						name: 'Thandiwe Mbeki',
						title: 'Instrument engineer',
						location: 'Greywharf',
						started: '2022-05-23',
					},
				],
			},
		],
	},
	{
		name: 'Marine Energy',
		hue: 'teal',
		head: {
			name: 'Idris Okonjo',
			title: 'Head of Marine Energy',
			location: 'Greywharf',
			started: '2022-01-10',
		},
		teams: [
			{
				lead: {
					name: 'Felix Brandauer',
					title: 'Lead, Tidal Turbines',
					location: 'Greywharf',
					started: '2022-03-28',
				},
				members: [
					{
						name: 'Oskar Lindqvist',
						title: 'Turbine design engineer',
						location: 'Greywharf',
						started: '2022-11-14',
					},
					{
						name: 'Yuki Tanabe',
						title: 'Structural analyst',
						location: 'Port Averly',
						started: '2023-04-03',
						status: 'Seconded',
					},
				],
			},
			{
				lead: {
					name: 'Noor Haddad',
					title: 'Lead, Grid and Storage',
					location: 'Port Averly',
					started: '2023-02-06',
				},
				members: [
					{
						name: 'Beatriz Almada',
						title: 'Power systems engineer',
						location: 'Port Averly',
						started: '2024-06-17',
					},
				],
			},
		],
	},
	{
		name: 'Institute Services',
		hue: 'plum',
		head: {
			name: 'Leona Castellanos',
			title: 'Head of Institute Services',
			location: 'Port Averly',
			started: '2019-09-16',
		},
		teams: [
			{
				lead: {
					name: 'Calla Whitlow',
					title: 'Lead, Research Computing',
					location: 'Port Averly',
					started: '2020-10-05',
				},
				members: [
					{
						name: 'Cormac Delaney',
						title: 'Systems administrator',
						location: 'Port Averly',
						started: '2021-01-11',
					},
					{
						name: 'Zofia Marchetti',
						title: 'Research software engineer',
						location: 'Port Averly',
						started: '2025-03-03',
						fte: 0.6,
					},
				],
			},
			{
				lead: {
					name: 'Dmitri Volkonsky',
					title: 'Lead, Finance and Grants',
					location: 'Port Averly',
					started: '2020-02-24',
				},
				members: [
					{
						name: 'Ilse Brandvold',
						title: 'Grants officer',
						location: 'Port Averly',
						started: '2024-09-09',
					},
				],
			},
		],
	},
];

/** A work address on the reserved `.example` domain. */
export function email(post: Post): string {
	if (post.status === 'Vacant') return '';
	const [first = '', ...rest] = post.name.toLowerCase().split(' ');
	return `${first.slice(0, 1)}.${rest.join('')}@saltmarsh.example`;
}
