import type {
	Bibliography,
	BibliographyAuthor,
	BibliographySource,
	BibliographyStyle,
	CitationOptions,
	SourceField,
	WordReferences,
} from 'ooxml-core/docx/ui';
import type { ReferencesText } from './strings';
import type { ReferencesDialog } from './dialog';
import { citationOptions } from './entry-dialogs';

const TYPE_LABELS = [
	'Book',
	'Book section',
	'Journal article',
	'Web page',
	'Report',
] as const satisfies readonly ReferencesText[];
const FIELD_LABELS: Record<SourceField, ReferencesText> = {
	Title: 'Title',
	BookTitle: 'Book title',
	JournalName: 'Journal name',
	InternetSiteTitle: 'Website name',
	Publisher: 'Publisher',
	City: 'City',
	Year: 'Year',
	Month: 'Month',
	Day: 'Day',
	Volume: 'Volume',
	Issue: 'Issue',
	Pages: 'Pages',
	Edition: 'Edition',
	URL: 'URL',
	StandardNumber: 'Standard number',
};

const authorText = (author: BibliographyAuthor | undefined) =>
	author?.kind === 'corporate'
		? author.name
		: (author?.people ?? [])
				.map(({ last, first, middle }) =>
					[last, [first, middle].filter(Boolean).join(' ')].filter(Boolean).join(', '),
				)
				.join('; ');

function authorFrom(text: string, corporate: boolean): BibliographyAuthor | undefined {
	if (!text) return undefined;
	if (corporate) return { kind: 'corporate', name: text };
	return {
		kind: 'people',
		people: text
			.split(';')
			.map((author) => author.trim())
			.filter(Boolean)
			.map((author) => {
				const comma = author.indexOf(',');
				const words = author.split(/\s+/u);
				const last = comma < 0 ? (words.pop() ?? '') : author.slice(0, comma).trim();
				const [first = '', ...middle] =
					comma < 0
						? words
						: author
								.slice(comma + 1)
								.trim()
								.split(/\s+/u);
				return { first, middle: middle.join(' '), last };
			}),
	};
}

const sourceChoices = (data: Bibliography): [string, string][] =>
	data.sources.map((source) => [
		source.tag,
		source.fields.Title ? `${source.fields.Title} (${source.tag})` : source.tag,
	]);

function checkTag(dialog: ReferencesDialog, tag: string, data: Bibliography, limit: number) {
	if (!tag || tag.length > 64 || !/^[\p{L}\p{N}_.-]+$/u.test(tag))
		throw new Error(
			dialog.t('Enter a source tag using letters, numbers, periods, hyphens, or underscores'),
		);
	if (data.sources.some((source) => source.tag.toLowerCase() === tag.toLowerCase()))
		throw new Error(dialog.t('A source already uses this tag'));
	if (data.sources.length >= limit) throw new Error(dialog.t('The source list is full'));
}

export function manageSourcesDialog(
	{ SOURCE_FIELDS, SOURCE_LIMIT, SOURCE_TYPES }: WordReferences,
	dialog: ReferencesDialog,
	data: Bibliography,
	save: (sources: BibliographySource[]) => void,
) {
	const existing = dialog.select('Source', [['', dialog.t('New source')], ...sourceChoices(data)]);
	const tag = dialog.text('Tag');
	tag.maxLength = 64;
	const types = SOURCE_TYPES.map((kind, index): [string, string] => [
		kind,
		dialog.t(TYPE_LABELS[index] ?? 'Book'),
	]);
	const known = new Set(types.map(([kind]) => kind));
	for (const source of data.sources)
		if (source.type && !known.has(source.type)) {
			types.push([source.type, source.type]);
			known.add(source.type);
		}
	const type = dialog.select('Source type', types);
	const corporate = dialog.check('Corporate author');
	const author = dialog.text('Author');
	author.placeholder = dialog.t('Last, First; Last, First');
	const fieldInputs = SOURCE_FIELDS.map((field) => ({
		field,
		input: dialog.text(FIELD_LABELS[field]),
	}));
	const current = () => data.sources.find((source) => source.tag === existing.value());
	const remove = dialog.action('Remove Source', () =>
		save(data.sources.filter((source) => source !== current())),
	);
	const sync = () => {
		const source = current();
		tag.value = source?.tag ?? '';
		tag.readOnly = Boolean(source);
		type.input.value = source?.type || 'Book';
		corporate.checked = source?.author?.kind === 'corporate';
		author.value = authorText(source?.author);
		for (const { field, input } of fieldInputs) input.value = source?.fields[field] ?? '';
		remove.disabled = !source;
		remove.title = source ? '' : dialog.t('Select an existing source');
	};
	existing.input.addEventListener('change', sync);
	sync();
	dialog.show(
		remove,
		dialog.action(
			'Save Source',
			() => {
				const before = current();
				const key = tag.value.trim();
				if (!before) checkTag(dialog, key, data, SOURCE_LIMIT);
				const fields: BibliographySource['fields'] = {};
				for (const { field, input } of fieldInputs) {
					const value = input.value.trim();
					if (value) fields[field] = value;
				}
				if (!fields.Title) throw new Error(dialog.t('Enter a title for this source'));
				const by =
					before &&
					author.value === authorText(before.author) &&
					corporate.checked === (before.author?.kind === 'corporate')
						? before.author
						: authorFrom(author.value.trim(), corporate.checked);
				const next = {
					tag: key,
					type: type.value() ?? 'Book',
					fields,
					...(by ? { author: by } : {}),
				};
				save(
					before
						? data.sources.map((source) => (source === before ? next : source))
						: [...data.sources, next],
				);
			},
			true,
		),
	);
}

export function insertCitationDialog(
	{ SOURCE_LIMIT }: WordReferences,
	dialog: ReferencesDialog,
	data: Bibliography,
	insert: (tag: string, options: CitationOptions, placeholder?: BibliographySource) => void,
) {
	const source = dialog.select('Source', [
		['', dialog.t('Add placeholder')],
		...sourceChoices(data),
	]);
	if (data.sources[0]) source.input.value = data.sources[0].tag;
	const tag = dialog.text('Placeholder tag');
	tag.maxLength = 64;
	const options = citationOptions(dialog);
	const sync = () => {
		tag.disabled = Boolean(source.value());
		tag.title = tag.disabled ? dialog.t('Choose Add placeholder to enter a new tag') : '';
	};
	source.input.addEventListener('change', sync);
	sync();
	dialog.show(
		dialog.action(
			'Insert',
			() => {
				const existing = source.value();
				if (existing) insert(existing, options());
				else {
					const key = tag.value.trim();
					checkTag(dialog, key, data, SOURCE_LIMIT);
					insert(key, options(), { tag: key, type: 'Book', fields: {}, placeholder: true });
				}
			},
			true,
		),
	);
}

export function styleDialog(
	dialog: ReferencesDialog,
	current: BibliographyStyle | undefined,
	apply: (style: BibliographyStyle) => void,
) {
	const style = dialog.select(
		'Citation Style',
		[
			['APA', 'APA'],
			['MLA', 'MLA'],
			['Chicago', dialog.t('Chicago author-date')],
		],
		current ?? 'APA',
	);
	dialog.show(dialog.action('OK', () => apply(style.value() ?? 'APA'), true));
}
