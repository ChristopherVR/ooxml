import { $, apps, badge } from './ui.js';
export const product = document.body.dataset.product || '';
export const suiteBase = new URL(
	document.querySelector('meta[name="ooxml-root"]')?.content || './',
	location.href,
);
export const productPaths = {
	docx: 'word',
	xlsx: 'excel',
	pptx: 'powerpoint',
	vsdx: 'visio',
	teams: 'teams',
};
export function workspaceUrl(type, id) {
	const url = new URL(
		/^https?:$/.test(suiteBase.protocol) ? suiteBase : 'https://ooxml.local.invalid/',
	);
	url.hash = `/${type}/${id}`;
	return url.href;
}
export function workspaceId(raw, type) {
	try {
		const url = new URL(raw);
		const base = new URL(workspaceUrl(type, 'base'));
		const paths = [
			base.pathname,
			...Object.values(productPaths).map((p) => new URL(`apps/${p}/`, suiteBase).pathname),
		];
		return url.origin === base.origin &&
			paths.includes(url.pathname) &&
			new RegExp(`^#/${type}/[\\w-]+$`, 'u').test(url.hash)
			? url.hash.slice(type.length + 3)
			: null;
	} catch {
		return null;
	}
}
export function mountProduct() {
	if (!product) {
		const links = document.createElement('details');
		links.className = 'standalone-links';
		links.innerHTML = `<summary>Open apps separately</summary>${Object.entries(productPaths)
			.map(
				([id, path]) =>
					`<a href="${new URL(`apps/${path}/`, suiteBase)}">${apps.find((a) => a.id === id)?.name ?? 'Teams'}</a>`,
			)
			.join('')}`;
		$('app-launcher').append(links);
		return;
	}
	const app = apps.find((a) => a.id === product) ?? { name: 'Teams', id: 'teams' };
	const brand = document.querySelector('.brand');
	brand.innerHTML = `${badge(product)}<strong>${app.name}</strong>`;
	brand.setAttribute('aria-label', `${app.name} home`);
	for (const button of $('rail').querySelectorAll('[data-app], [data-teams]'))
		button.hidden = (button.dataset.app || 'teams') !== product;
	for (const button of $('create-row').querySelectorAll('[data-create]'))
		button.hidden = button.dataset.create !== product;
	$('app-launcher').innerHTML =
		`<div class="launcher-heading"><h2>Your apps</h2></div><div class="launcher-grid"><a href="${new URL('?suite', suiteBase)}">OOXML Office</a>${[...apps, { id: 'teams', name: 'Teams' }].map((a) => `<a href="${new URL(`apps/${productPaths[a.id]}/`, suiteBase)}">${badge(a.id)}<span>${a.name}</span></a>`).join('')}</div>`;
}
