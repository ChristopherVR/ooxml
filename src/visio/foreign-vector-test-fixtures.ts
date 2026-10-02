export const path = (d = 'M0 0L10 0L10 10Z', attrs = {}) => ({
	tag: 'path',
	attrs: { d, ...attrs },
});
export const group = (children: unknown[], attrs = {}) => ({ tag: 'g', attrs, children });
export const clip = (id: string, children: unknown[] = [path()], attrs = {}) => ({
	tag: 'clipPath',
	attrs: { id, ...attrs },
	children,
});
export const defs = (children: unknown[]) => ({ tag: 'defs', attrs: {}, children });
export const root = (children: unknown[] = [path()], attrs = {}) => ({
	tag: 'svg',
	attrs: {
		xmlns: 'http://www.w3.org/2000/svg',
		width: 20,
		height: 20,
		viewBox: '0 0 20 20',
		style: 'isolation:isolate',
		...attrs,
	},
	children,
});
