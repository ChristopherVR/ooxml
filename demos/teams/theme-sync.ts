/**
 * Follows the docs site's appearance choice. VitePress stores it in localStorage under
 * `vitepress-theme-appearance` ('light', 'dark' or 'auto'); every viewer demo shares that key
 * because they are all served from the same origin. The office-ui tokens read
 * `data-office-theme` on the root, so setting it restyles the running app without a reload.
 * With no stored choice (or 'auto') the demo keeps following the operating system.
 */
const KEY = 'vitepress-theme-appearance';

function apply(value: string | null): void {
	const root = document.documentElement;
	if (value === 'light' || value === 'dark') root.setAttribute('data-office-theme', value);
	else root.removeAttribute('data-office-theme');
}

function stored(): string | null {
	try {
		return localStorage.getItem(KEY);
	} catch {
		return null;
	}
}

apply(stored());
window.addEventListener('storage', (event) => {
	if (event.key === KEY || event.key === null) apply(event.newValue);
});

export {};
