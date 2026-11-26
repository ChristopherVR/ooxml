import type {
	CitationOptions,
	IndexLeader,
	IndexOptions,
	WordReferences,
} from 'ooxml-core/docx/ui';
import { labelled, numberInput } from '../dialog-fields';
import type { ReferencesText } from './strings';
import type { ReferencesDialog } from './dialog';

type MarkOptions = {
	main: string;
	subentry: string;
	crossReference: string;
	bookmark: string;
	bold: boolean;
	italic: boolean;
};

export function markIndexDialog(
	dialog: ReferencesDialog,
	selection: string,
	bookmarks: string[],
	apply: (options: MarkOptions, all: boolean) => void,
) {
	const main = dialog.text('Main entry', selection);
	const subentry = dialog.text('Subentry');
	const crossReference = dialog.text('Cross-reference');
	const bookmark = dialog.select('Page range bookmark', [
		['', dialog.t('None')],
		...bookmarks.map((name): [string, string] => [name, name]),
	]);
	const bold = dialog.check('Bold page numbers');
	const italic = dialog.check('Italic page numbers');
	const run = (all: boolean) =>
		apply(
			{
				main: main.value.trim(),
				subentry: subentry.value.trim(),
				crossReference: crossReference.value.trim(),
				bookmark: bookmark.value() ?? '',
				bold: bold.checked,
				italic: italic.checked,
			},
			all,
		);
	const mark = dialog.action('Mark', () => run(false), true);
	const markAll = dialog.action('Mark All', () => run(true));
	const sync = () => {
		mark.disabled = !main.value.trim();
		markAll.disabled =
			!selection || mark.disabled || Boolean(bookmark.value() || crossReference.value.trim());
		mark.title = mark.disabled ? dialog.t('Enter a main entry') : '';
		markAll.title = markAll.disabled ? dialog.t('Select text to mark all page references') : '';
	};
	dialog.content.addEventListener('input', sync);
	sync();
	dialog.show(markAll, mark);
}

const LEADERS: Record<IndexLeader, ReferencesText> = {
	none: 'None',
	dot: 'Dot',
	hyphen: 'Dash',
	underscore: 'Underline',
};

export function insertIndexDialog(
	{ INDEX_LEADERS, REFERENCE_LCID }: WordReferences,
	dialog: ReferencesDialog,
	locale: string,
	apply: (options: IndexOptions) => void,
) {
	const columns = numberInput(1, 4, 1);
	columns.value = '1';
	dialog.content.append(labelled(dialog.t('Columns'), columns));
	const align = dialog.check('Right-align page numbers', true);
	const leader = dialog.select(
		'Tab leader',
		INDEX_LEADERS.map((key) => [key, dialog.t(LEADERS[key])]),
		'dot',
	);
	const layout = dialog.select('Index layout', [
		['indented', dialog.t('Indented')],
		['runIn', dialog.t('Run-in')],
	]);
	const languages = new Intl.DisplayNames([locale], { type: 'language' });
	const language = dialog.select(
		'Language',
		Object.keys(REFERENCE_LCID).map((tag) => [tag, languages.of(tag) ?? tag]),
		locale,
	);
	align.addEventListener('change', () => {
		leader.input.disabled = !align.checked;
		leader.input.title = align.checked ? '' : dialog.t('Right-align page numbers');
	});
	dialog.show(
		dialog.action(
			'OK',
			() => {
				if (!columns.checkValidity() || !columns.value)
					throw new Error(dialog.t('Choose between one and four columns'));
				apply({
					columns: columns.valueAsNumber,
					runIn: layout.value() === 'runIn',
					rightAlign: align.checked,
					leader: leader.value() ?? 'dot',
					lcid: REFERENCE_LCID[language.value() ?? locale] ?? 1033,
				});
			},
			true,
		),
	);
}

function authorityCategory(references: WordReferences, dialog: ReferencesDialog) {
	return dialog.select(
		'Category',
		Array.from({ length: 16 }, (_, index) => {
			const label = references.AUTHORITY_CATEGORIES[index];
			return [String(index + 1), label ? dialog.t(label) : String(index + 1)] as const;
		}),
	);
}

export function markAuthorityDialog(
	references: WordReferences,
	dialog: ReferencesDialog,
	selection: string,
	apply: (long: string, short: string, category: number) => void,
) {
	const long = dialog.text('Long citation', selection);
	const short = dialog.text('Short citation', selection);
	const category = authorityCategory(references, dialog);
	const mark = dialog.action(
		'Mark',
		() =>
			apply(long.value.trim(), short.value.trim() || long.value.trim(), Number(category.value())),
		true,
	);
	const sync = () => {
		mark.disabled = !long.value.trim();
		mark.title = mark.disabled ? dialog.t('Enter a long citation') : '';
	};
	long.addEventListener('input', sync);
	sync();
	dialog.show(mark);
}

export function authorityDialog(
	references: WordReferences,
	dialog: ReferencesDialog,
	apply: (category: number, passim: boolean, keep: boolean) => void,
) {
	const category = authorityCategory(references, dialog);
	const passim = dialog.check('Use passim', true);
	const keep = dialog.check('Keep original formatting', true);
	dialog.show(
		dialog.action('OK', () => apply(Number(category.value()), passim.checked, keep.checked), true),
	);
}

export function citationOptions(dialog: ReferencesDialog, current?: CitationOptions) {
	const pages = dialog.text('Pages', current?.pages);
	const suppressAuthor = dialog.check('Suppress author', current?.suppressAuthor);
	const suppressTitle = dialog.check('Suppress title', current?.suppressTitle);
	const suppressYear = dialog.check('Suppress year', current?.suppressYear);
	return (): CitationOptions => ({
		pages: pages.value.trim(),
		suppressAuthor: suppressAuthor.checked,
		suppressTitle: suppressTitle.checked,
		suppressYear: suppressYear.checked,
	});
}
