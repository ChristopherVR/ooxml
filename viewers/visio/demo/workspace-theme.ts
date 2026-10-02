/** The site and workspace share one optional, local-only appearance preference. */
export function wireWorkspaceTheme(doc: Document): () => void {
	const view = doc.defaultView;
	if (!view) return () => {};
	const root = doc.documentElement;
	const button = doc.getElementById('theme-toggle');
	const storageKey = 'visio-docs-theme';
	if (new URL(view.location.href).searchParams.get('embed') === '1') root.dataset.embedded = '';
	function syncLabel(): void {
		const next = root.dataset.theme === 'light' ? 'dark' : 'light';
		doc
			.querySelector('meta[name="theme-color"]')
			?.setAttribute('content', next === 'dark' ? '#ffffff' : '#171a1e');
		button?.setAttribute('aria-label', `Switch to ${next} theme`);
		button?.setAttribute('title', `Switch to ${next} theme`);
		if (button) button.textContent = next === 'light' ? '☀' : '◐';
	}
	try {
		const saved = view.localStorage.getItem(storageKey);
		if (saved === 'light' || saved === 'dark') root.dataset.theme = saved;
	} catch {
		// The current tab still has a working theme when storage is disabled.
	}
	syncLabel();
	const observer = new view.MutationObserver(syncLabel);
	observer.observe(root, { attributes: true, attributeFilter: ['data-theme'] });
	const onTheme = (): void => {
		root.dataset.theme = root.dataset.theme === 'light' ? 'dark' : 'light';
		try {
			view.localStorage.setItem(storageKey, root.dataset.theme);
		} catch {
			// Keep the appearance choice in this tab.
		}
		syncLabel();
	};
	const onStorage = (event: StorageEvent): void => {
		if (event.key === storageKey && (event.newValue === 'light' || event.newValue === 'dark')) {
			root.dataset.theme = event.newValue;
			syncLabel();
		}
	};
	button?.addEventListener('click', onTheme);
	view.addEventListener('storage', onStorage);
	return () => {
		observer.disconnect();
		button?.removeEventListener('click', onTheme);
		view.removeEventListener('storage', onStorage);
	};
}
