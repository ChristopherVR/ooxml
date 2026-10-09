import { $, apps, badge, escape } from './ui.js';
import { currentProfile } from './profile-state.js';
import { icon } from './icons.js';

const key = () => `suite-favourites-${currentProfile().id}`;
function favourites() {
	try {
		const items = JSON.parse(localStorage.getItem(key()));
		return new Set(Array.isArray(items) ? items : []);
	} catch {
		return new Set();
	}
}
export function toggleFavourite(id) {
	const items = favourites();
	if (items.has(id)) {
		items.delete(id);
	} else {
		items.add(id);
	}
	localStorage.setItem(key(), JSON.stringify([...items]));
}
export function renderLibrary(files, filter) {
	const query = $('search').value.toLowerCase();
	const saved = favourites();
	const rows = files
		.filter(
			(d) =>
				!d.parent &&
				(!document.body.dataset.product ||
					document.body.dataset.product === 'teams' ||
					d.kind === document.body.dataset.product) &&
				(filter === 'favourites'
					? saved.has(d.id)
					: !filter || filter === 'files' || d.kind === filter) &&
				d.name.toLowerCase().includes(query),
		)
		.sort((a, b) => b.modified - a.modified);
	const profile = currentProfile();
	const app = apps.find((a) => a.id === filter);
	$('home-title').textContent =
		filter === 'favourites'
			? 'Favourites'
			: filter === 'files'
				? 'My files'
				: app
					? app.name
					: 'Home';
	// An app's own view keeps its create tile; the other library views show files only.
	const product = document.body.dataset.product;
	for (const button of $('create-row').querySelectorAll('[data-create]'))
		button.hidden =
			(!!product && button.dataset.create !== product) ||
			(!!app && button.dataset.create !== app.id);
	$('create-section').hidden =
		(!!filter && !app) || !!query || !$('create-row').querySelector('[data-create]:not([hidden])');
	$('welcome').hidden = !!filter || !!query;
	$('library-title').textContent = query
		? `Results for “${$('search').value}”`
		: filter === 'favourites'
			? 'Saved for quick access'
			: 'Your documents';
	$('files').innerHTML = rows
		.map(
			(d) =>
				`<tr><td><button class="file-name" data-open="${d.id}">${badge(d.kind)}<span><strong>${escape(d.name)}</strong><small>${escape(profile.organization)}</small></span></button></td><td>${escape(apps.find((a) => a.id === d.kind)?.name)}</td><td><span class="location-label">${icon('device')}App storage</span></td><td>${new Date(d.modified).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}</td><td class="file-actions"><button data-favourite="${d.id}" aria-label="${saved.has(d.id) ? 'Remove from' : 'Add to'} favourites: ${escape(d.name)}" aria-pressed="${saved.has(d.id)}">${icon('star')}</button><button data-download="${d.id}" aria-label="Download ${escape(d.name)}">↓</button><button data-file-menu="${d.id}" aria-label="File actions: ${escape(d.name)}">⋯</button></td></tr>`,
		)
		.join('');
	$('empty').hidden = rows.length > 0;
	$('empty-title').textContent = query
		? 'No matching files'
		: filter === 'favourites'
			? 'Keep your important files close'
			: 'Start with your first document';
	$('empty-description').textContent = query
		? 'Try another file name or choose a different application.'
		: filter === 'favourites'
			? 'Select the star beside a file to find it here.'
			: 'Create something new above, or bring in a file from your device.';
	$('file-count').textContent = `${rows.length} file${rows.length === 1 ? '' : 's'}`;
	$('library-owner').textContent = profile.organization;
	for (const button of document.querySelectorAll('[data-library]'))
		button.setAttribute(
			'aria-current',
			String(document.body.dataset.view === 'home' && button.dataset.library === filter),
		);
	for (const button of document.querySelectorAll('[data-file-filter]'))
		button.setAttribute('aria-pressed', String(button.dataset.fileFilter === (app ? filter : '')));
	const bytes = files.reduce((sum, d) => sum + d.bytes.length, 0);
	$('storage-summary').textContent =
		`${files.filter((d) => !d.parent).length} files · ${bytes < 1024 * 1024 ? `${Math.ceil(bytes / 1024)} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`}`;
}
