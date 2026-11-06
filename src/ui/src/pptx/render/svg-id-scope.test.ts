// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';

import { installSvgIdScope, scopeSvgIds } from './svg-id-scope';

/** One painted copy of a shape: a silhouette filter, a marker and a gradient. */
function copy(host: HTMLElement, hidden: boolean): HTMLElement {
	const el = document.createElement('div');
	el.setAttribute('data-element-id', 's1');
	if (hidden) {
		el.style.display = 'none';
	}
	el.innerHTML =
		'<svg><defs><filter id="shape-fx-s1"><feFlood/></filter>' +
		'<marker id="s1-end"><path d="M0 0"/></marker>' +
		'<linearGradient id="s1-g"/></defs>' +
		'<path d="M0 0" marker-end="url(#s1-end)" fill="url(#s1-g)"/></svg>' +
		'<div style="filter: url(#shape-fx-s1)"></div>';
	host.append(el);
	return el;
}

function resolves(root: Element): boolean {
	return [...root.querySelectorAll('[marker-end],[fill],[style]')]
		.flatMap((n) =>
			['marker-end', 'fill', 'style'].map((a) => /#([^)"']+)/u.exec(n.getAttribute(a) ?? '')?.[1]),
		)
		.filter((id): id is string => id !== undefined)
		.every((id) => root.querySelector(`[id="${id}"]`) !== null);
}

describe('scopeSvgIds', () => {
	afterEach(() => {
		document.body.innerHTML = '';
	});

	it('gives each copy its own def ids and points references at their own copy', () => {
		const first = copy(document.body, true);
		const second = copy(document.body, false);
		scopeSvgIds(document.body);
		const ids = [...document.body.querySelectorAll('[id]')].map((n) => n.id);
		expect(new Set(ids).size).toBe(ids.length);
		expect(ids).toHaveLength(6);
		expect(resolves(first)).toBe(true);
		expect(resolves(second)).toBe(true);
	});

	it('is idempotent and keeps the original id in data-pptx-sid', () => {
		const only = copy(document.body, false);
		scopeSvgIds(document.body);
		const before = only.innerHTML;
		scopeSvgIds(document.body);
		expect(only.innerHTML).toBe(before);
		expect(only.querySelector('filter')?.getAttribute('data-pptx-sid')).toBe('shape-fx-s1');
	});

	it('scopes copies mounted later and clones of scoped copies', async () => {
		const release = installSvgIdScope(document);
		try {
			const first = copy(document.body, false);
			await Promise.resolve();
			const clone = first.cloneNode(true) as HTMLElement;
			document.body.append(clone);
			await Promise.resolve();
			const ids = [...document.body.querySelectorAll('[id]')].map((n) => n.id);
			expect(new Set(ids).size).toBe(ids.length);
			expect(resolves(first)).toBe(true);
			expect(resolves(clone)).toBe(true);
		} finally {
			release();
		}
	});
});
