import { dialogButton, labelled, textInput } from './dialog-fields';
import {
	allowedVideoResults,
	formatVideoDuration,
	MAX_VIDEO_QUERY_LENGTH,
	type OnlineVideoSearchProvider,
	type OnlineVideoSearchResult,
} from './online-video-search';
import {
	localizeElement,
	translate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';

const t = (locale: EditorLocale, key: string) => translate(locale, key as LocalizationKey);

export interface VideoSearchPanel {
	element: HTMLElement;
	/** Shows the panel for `provider` (hidden without one) and clears earlier results. */
	reset(provider: OnlineVideoSearchProvider | undefined, locale: EditorLocale): void;
	/** Stops a search in flight (the dialog closed). */
	cancel(): void;
}

/**
 * Insert > Online Video's search box, shown only when the host gives a video search provider.
 * Results the allow-list does not accept are never listed; choosing one fills the address box.
 */
export function createVideoSearchPanel(
	hosts: () => readonly string[],
	choose: (result: OnlineVideoSearchResult) => void,
): VideoSearchPanel {
	const element = document.createElement('section');
	element.className = 'dve-online-video-search';
	const query = textInput();
	query.type = 'search';
	query.maxLength = MAX_VIDEO_QUERY_LENGTH;
	const search = dialogButton('Search');
	const status = document.createElement('p');
	status.setAttribute('role', 'status');
	const list = document.createElement('ul');
	list.setAttribute('role', 'listbox');
	list.style.cssText = 'list-style:none;margin:0;padding:0;max-height:16em;overflow:auto';
	const label = labelled('Search', query);
	const row = document.createElement('div');
	row.style.cssText = 'display:flex;gap:.5em;align-items:end';
	row.append(label, search);
	element.append(row, status, list);
	let provider: OnlineVideoSearchProvider | undefined;
	let locale: EditorLocale = 'en';
	let pending: AbortController | undefined;

	const show = (results: OnlineVideoSearchResult[], found: number) => {
		list.replaceChildren(
			...results.map((result) => {
				const item = document.createElement('li');
				const button = document.createElement('button');
				button.type = 'button';
				button.setAttribute('role', 'option');
				button.style.cssText =
					'display:flex;gap:.5em;width:100%;text-align:start;align-items:center;padding:.25em';
				if (result.thumbnailUrl) {
					const image = document.createElement('img');
					image.src = result.thumbnailUrl;
					image.alt = '';
					image.width = 96;
					image.height = 54;
					image.loading = 'lazy';
					image.referrerPolicy = 'no-referrer';
					button.append(image);
				}
				const text = document.createElement('span');
				const title = document.createElement('strong');
				title.textContent = result.title;
				const meta = document.createElement('small');
				meta.style.display = 'block';
				meta.textContent = [
					result.publisher,
					result.durationSeconds === undefined
						? undefined
						: formatVideoDuration(result.durationSeconds),
					new URL(result.url).hostname,
				]
					.filter(Boolean)
					.join(' · ');
				text.append(title, meta);
				button.append(text);
				button.addEventListener('click', () => choose(result));
				item.append(button);
				return item;
			}),
		);
		status.textContent = results.length
			? found > results.length
				? t(locale, 'Some results are hidden because their site is not allowed.')
				: ''
			: t(locale, found ? 'No results come from an allowed site.' : 'No videos were found.');
	};

	const run = async () => {
		const words = query.value.trim().slice(0, MAX_VIDEO_QUERY_LENGTH);
		if (!provider || !words) return;
		pending?.abort();
		const controller = new AbortController();
		pending = controller;
		status.textContent = t(locale, 'Searching…');
		list.replaceChildren();
		search.disabled = true;
		try {
			const results = await provider.search(words, { signal: controller.signal, locale });
			if (pending !== controller) return;
			const found = Array.isArray(results) ? results : [];
			show(allowedVideoResults(found, hosts()), found.length);
		} catch {
			if (pending === controller)
				status.textContent = t(locale, 'The video search is not available right now.');
		} finally {
			if (pending === controller) {
				pending = undefined;
				search.disabled = false;
			}
		}
	};
	search.addEventListener('click', () => void run());
	query.addEventListener('keydown', (event) => {
		if (event.key !== 'Enter') return;
		// Enter searches here; it must not insert whatever the address box holds.
		event.preventDefault();
		event.stopPropagation();
		void run();
	});

	return {
		element,
		reset(next, nextLocale) {
			pending?.abort();
			pending = undefined;
			provider = next;
			locale = nextLocale;
			element.hidden = !next;
			query.value = '';
			search.disabled = false;
			status.textContent = '';
			list.replaceChildren();
			label.querySelector('span')!.textContent = next
				? `${t(locale, 'Search')} ${next.name}`.slice(0, 80)
				: '';
			localizeElement(search, locale);
		},
		cancel() {
			pending?.abort();
			pending = undefined;
		},
	};
}
