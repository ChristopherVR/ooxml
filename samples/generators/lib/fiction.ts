/**
 * The invented world the samples share. Every company, product, place and person here is
 * fictitious; none stands for a real organisation, brand or individual.
 */

/** Open fonts only (docs/font-loading.md): all three are served by the webfont fallback. */
export const FONTS = {
	/** Headings and display text. */
	display: 'Rubik',
	/** Body text in decks, sheets and diagrams. */
	body: 'Noto Sans',
	/** Long-form reading text. */
	serif: 'Noto Serif',
} as const;

export const PEOPLE = [
	'Mara Ellingwood',
	'Tobias Renner',
	'Priya Vasanth',
	'Idris Okonjo',
	'Leona Castellanos',
	'Hugo Falkenrath',
	'Sunniva Dahlgren',
	'Kenji Morikawa',
	'Amara Diallo',
	'Felix Brandauer',
	'Noor Haddad',
	'Calla Whitlow',
	'Dmitri Volkonsky',
	'Anouk Verhaegen',
	'Rafael Monteverde',
	'Esi Agyeman',
	'Linnea Sorensen',
	'Marek Zielonka',
	'Thandiwe Mbeki',
	'Oskar Lindqvist',
	'Yuki Tanabe',
	'Beatriz Almada',
	'Cormac Delaney',
	'Zofia Marchetti',
] as const;

export const REGIONS = [
	'Northwind Coast',
	'Central Plains',
	'Highland',
	'Lakeshore',
	'Southern Delta',
] as const;

export const CITIES = [
	'Port Averly',
	'Kessington',
	'Marrow Bay',
	'Elden Ford',
	'Stonehaven Mills',
	'Larkspur Heights',
	'Quillhaven',
	'Brackwater',
	'Sunderby',
	'Vale Orrin',
	'Thistledown',
	'Greywharf',
] as const;
