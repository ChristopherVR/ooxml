/**
 * The site's one not-found page (GitHub Pages serves it for every missing path under the site).
 * The docs sites translate only part of their pages, so a language-menu link from an untranslated
 * page points at /<product>/<locale>/<page> that does not exist: send it to the English page.
 * Any other missing address shows the page.
 */
const root = new URL('./', document.currentScript?.src ?? import.meta.url).pathname;
const match = new RegExp(`^${root}(docx|xlsx|pptx|visio|teams)/(fr|es|de|zh)/(.*)$`).exec(
	location.pathname,
);
if (match) {
	const target = `${root}${match[1]}/${match[3]}${location.search}${location.hash}`;
	document.querySelector('#notfound-message').textContent =
		'This page is not translated yet. Taking you to the English version.';
	location.replace(target);
}
