const idPattern = /^WA\d{1,16}$/i;

/** Extract an Office Marketplace asset identifier without following the link. */
export function marketplaceAssetId(input: string): string | undefined {
	const value = input.trim();
	if (idPattern.test(value)) return value.toUpperCase();
	if (value.length > 2048 || !/^https:\/\//i.test(value) || /[\\\s]/.test(value)) return;
	let url: URL;
	try {
		url = new URL(value);
	} catch {
		return;
	}
	if (url.username || url.password || url.port || value.split('/')[2]?.includes('@')) return;
	const tail = url.pathname.split('/').filter(Boolean).at(-1) ?? '';
	const ids = (() => {
		switch (url.hostname) {
			case 'appsource.microsoft.com':
			case 'marketplace.microsoft.com':
				return /\/(?:product|marketplace\/apps)\/office\//i.test(url.pathname) ? [tail] : [];
			case 'store.office.com':
				return /\/app\.aspx$/i.test(url.pathname) ? url.searchParams.getAll('assetid') : [];
			case 'go.microsoft.com':
				return /^\/fwlink\/?$/i.test(url.pathname) ? url.searchParams.getAll('templateid') : [];
			default:
				return [];
		}
	})();
	const [id] = ids;
	if (ids.length === 1 && id && idPattern.test(id)) return id.toUpperCase();
}
