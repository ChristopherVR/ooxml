const widths = new WeakMap<Document, Map<string, number>>();
/** CSS border widths may be snapped to device pixels; keep source content origins independent. */
export function measuredBorderWidth(css: string, fallback: number): number {
	if (!css || fallback === 0) return 0;
	const view = document.defaultView;
	if (!view) return fallback;
	let cache = widths.get(document);
	if (!cache) {
		cache = new Map();
		widths.set(document, cache);
	}
	const key = `${view.devicePixelRatio}|${css}`;
	const cached = cache.get(key);
	if (cached !== undefined) return cached;
	const probe = document.createElement('span');
	probe.dataset.printBorderProbe = 'true';
	probe.setAttribute('aria-hidden', 'true');
	Object.assign(probe.style, {
		position: 'fixed',
		visibility: 'hidden',
		pointerEvents: 'none',
		left: '-10000px',
		top: '0px',
		width: '0px',
		height: '0px',
		borderLeft: css,
	});
	const parent = document.body ?? document.documentElement;
	if (!parent) return fallback;
	parent.append(probe);
	try {
		const actual = Number.parseFloat(view.getComputedStyle(probe).borderLeftWidth);
		const width = Number.isFinite(actual) ? actual : fallback;
		cache.set(key, width);
		return width;
	} finally {
		probe.remove();
	}
}
