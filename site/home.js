import { APPS, appIcon } from './apps.js?v=launcher-20261007';

const SOURCE = 'https://github.com/ChristopherVR/ooxml/tree/main';
const arrow = '<span class="link-arrow" aria-hidden="true">↗</span>';
const tag = (app) =>
	app.tag ? `<span class="tag tag--${app.tag.tone}">${app.tag.label}</span>` : '';

/** The directory opens docs; launch controls keep the suite's existing tab routes. */
export function initHome() {
	for (const [id, apps] of [
		['app-grid', APPS.filter((app) => app.id !== 'teams')],
		['collaboration-grid', APPS.filter((app) => app.id === 'teams')],
	]) {
		document.getElementById(id).innerHTML = apps
			.map(
				(app) => `
			<li><a class="app" href="${app.docs}" style="--app:${app.color}">
				${appIcon(app)}<span class="app__name">${app.name}${tag(app)}</span>
				<span class="app__desc">${app.description}</span>
				<span class="app__link">Explore ${app.name} ${arrow}</span>
			</a></li>`,
			)
			.join('');
	}
	document.getElementById('launch-actions').innerHTML = APPS.map(
		(app) => `
		<a class="launch" href="#/${app.id}" style="--app:${app.color}">
			${appIcon(app)}<span>${app.id === 'teams' ? 'OpenTeams' : `Open ${app.name}`}</span>
			${tag(app)}<span class="launch__arrow" aria-hidden="true">→</span>
		</a>`,
	).join('');
	document.getElementById('source-groups').innerHTML = [
		[
			'Shared foundation',
			[
				['ooxml-core', `${SOURCE}/src/core`, 'Format models, parsing, editing and serialization.'],
				['ooxml-ui', `${SOURCE}/src/ui`, 'Shared browser controls and product editors.'],
			],
		],
		[
			'App bindings',
			APPS.map((app) => [
				`${app.repo}-viewer`,
				`${SOURCE}/viewers/${app.repo}`,
				`${app.name} framework bindings, demos and documentation.`,
			]),
		],
		[
			'Supporting libraries',
			[
				[
					'ole2',
					'https://github.com/ChristopherVR/ole2',
					'Legacy Office binary formats and compound-file containers.',
				],
				[
					'emf-converter',
					'https://github.com/ChristopherVR/emf-converter',
					'EMF and WMF graphics conversion.',
				],
				[
					'mtx-decompressor',
					'https://github.com/ChristopherVR/mtx-decompressor',
					'MicroType Express embedded font decompression.',
				],
			],
		],
	]
		.map(
			([title, links]) => `<section class="source-group">
		<h3>${title}</h3><dl class="core__list">${links
			.map(
				([name, href, description]) => `
			<div><dt><a href="${href}" target="_blank" rel="noreferrer">${name} ${arrow}</a></dt>
			<dd>${description}</dd></div>`,
			)
			.join('')}</dl>
		</section>`,
		)
		.join('');
}
