/**
 * The works the report cites, all invented. `sourcesXml` is Word's bibliography sources part
 * (`b:Sources`, ECMA-376 Part 1, 22.6); `entry` is the reference as Word's APA style renders it,
 * which is what the BIBLIOGRAPHY field stores as its result.
 */
import { xmlEscape } from '../../lib/package';

interface Source {
	tag: string;
	type: 'Report' | 'JournalArticle' | 'Book' | 'ConferenceProceedings';
	/** `[last, first]` pairs, or one corporate author. */
	authors: readonly (readonly [string, string])[] | string;
	year: number;
	/** Bibliography fields in the order Word writes them. */
	fields: Record<string, string>;
	/** The reference in `rich` markup. */
	entry: string;
}

export const SOURCES: readonly Source[] = [
	{
		tag: 'Cas24',
		type: 'Report',
		authors: [
			['Castellanos', 'Leona'],
			['Dahlgren', 'Sunniva'],
		],
		year: 2024,
		fields: {
			Title: 'Marine mammal monitoring protocol for small tidal arrays',
			Publisher: 'Northwind Coast Environment Board',
			City: 'Port Averly',
		},
		entry:
			'Castellanos, L., & Dahlgren, S. (2024). *Marine mammal monitoring protocol for small tidal arrays.* Port Averly: Northwind Coast Environment Board.',
	},
	{
		tag: 'Fal25',
		type: 'ConferenceProceedings',
		authors: [
			['Falkenrath', 'Hugo'],
			['Diallo', 'Amara'],
			['Whitlow', 'Calla'],
		],
		year: 2025,
		fields: {
			Title: 'Levelised cost of pilot-scale tidal arrays',
			ConferenceName: 'Proceedings of the Ninth Lakeshore Symposium on Ocean Energy',
			Pages: '55-63',
			City: 'Marrow Bay',
			Publisher: 'Lakeshore Ocean Energy Society',
		},
		entry:
			'Falkenrath, H., Diallo, A., & Whitlow, C. (2025). Levelised cost of pilot-scale tidal arrays. *Proceedings of the Ninth Lakeshore Symposium on Ocean Energy* (pp. 55-63). Marrow Bay: Lakeshore Ocean Energy Society.',
	},
	{
		tag: 'Mor21',
		type: 'Book',
		authors: [['Morikawa', 'Kenji']],
		year: 2021,
		fields: {
			Title: 'Hydrokinetic turbines: design and field practice',
			City: 'Kessington',
			Publisher: 'Brackwater Academic Press',
		},
		entry:
			'Morikawa, K. (2021). *Hydrokinetic turbines: design and field practice.* Kessington: Brackwater Academic Press.',
	},
	{
		tag: 'Nor24',
		type: 'Report',
		authors: 'Northwind Coast Grid Authority',
		year: 2024,
		fields: {
			Title: 'Connection code for embedded marine generation, issue 3',
			Publisher: 'Northwind Coast Grid Authority',
			City: 'Port Averly',
		},
		entry:
			'Northwind Coast Grid Authority. (2024). *Connection code for embedded marine generation, issue 3.* Port Averly: Northwind Coast Grid Authority.',
	},
	{
		tag: 'Ren23',
		type: 'Report',
		authors: [
			['Renner', 'Tobias'],
			['Okonjo', 'Idris'],
		],
		year: 2023,
		fields: {
			Title: 'Resource assessment of the Greywharf Narrows',
			Publisher: 'Saltmarsh Institute for Coastal Engineering',
			City: 'Greywharf',
		},
		entry:
			'Renner, T., & Okonjo, I. (2023). *Resource assessment of the Greywharf Narrows.* Greywharf: Saltmarsh Institute for Coastal Engineering.',
	},
	{
		tag: 'Vas22',
		type: 'JournalArticle',
		authors: [['Vasanth', 'Priya']],
		year: 2022,
		fields: {
			Title: 'Harmonic analysis of tidal currents in constricted channels',
			JournalName: 'Journal of Coastal Flow Studies',
			Pages: '101-118',
			Volume: '14',
			Issue: '2',
		},
		entry:
			'Vasanth, P. (2022). Harmonic analysis of tidal currents in constricted channels. *Journal of Coastal Flow Studies, 14*(2), 101-118.',
	},
];

const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';

export function sourcesXml(): string {
	const source = (item: Source, index: number) => {
		const authors =
			typeof item.authors === 'string'
				? `<b:Corporate>${xmlEscape(item.authors)}</b:Corporate>`
				: `<b:NameList>${item.authors
						.map(
							([last, first]) =>
								`<b:Person><b:Last>${xmlEscape(last)}</b:Last><b:First>${xmlEscape(first)}</b:First></b:Person>`,
						)
						.join('')}</b:NameList>`;
		const fields = Object.entries(item.fields)
			.map(([name, value]) => `<b:${name}>${xmlEscape(value)}</b:${name}>`)
			.join('');
		// A fixed GUID per source: Word expects one and nothing here may come from the runtime.
		const guid = `{5A17B0C0-0000-4000-8000-00000000000${index + 1}}`;
		return `<b:Source><b:Tag>${item.tag}</b:Tag><b:SourceType>${item.type}</b:SourceType><b:Guid>${guid}</b:Guid><b:Author><b:Author>${authors}</b:Author></b:Author>${fields}<b:Year>${item.year}</b:Year><b:RefOrder>${index + 1}</b:RefOrder></b:Source>`;
	};
	return `<?xml version="1.0" encoding="UTF-8" standalone="no"?><b:Sources xmlns:b="${NS}" xmlns="${NS}" SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6">${SOURCES.map(source).join('')}</b:Sources>`;
}
