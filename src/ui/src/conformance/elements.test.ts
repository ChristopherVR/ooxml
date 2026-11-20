// @vitest-fresh-modules: defines every element in a registry no earlier file has touched.
// The element conformance gate (`bun run check:elements`): every element this package defines,
// checked against the contract in src/ui/README.md ("How the elements are built"). Elements that
// fail a check today are listed in known-gaps.ts; the list may only shrink (see that file).
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { OfficeElement } from '../base';
import { CHECKS } from './contract';
import { type DefinedElement, defineEverything } from './elements';
import { KNOWN_GAPS } from './known-gaps';
import { trackResources } from './resources';

let elements: DefinedElement[] = [];
let errors: string[] = [];
const host = document.createElement('div');
// Tracking starts before anything is defined, so a listener an element adds on first use counts.
const resources = trackResources(window);

beforeAll(async () => {
	document.body.append(host);
	({ elements, errors } = await defineEverything(customElements));
}, 180_000);
afterAll(() => {
	resources.restore();
	host.remove();
});

describe('element conformance', () => {
	it('registers every element, and a second register call is a no-op', () => {
		expect(errors).toEqual([]);
		// Guards the discovery itself: an empty run would pass every check below.
		expect(elements.length).toBeGreaterThan(130);
	});

	for (const check of CHECKS)
		it(`every element ${check.rule} (${check.id})`, async () => {
			const failing: Record<string, string> = {};
			for (const element of elements) {
				if (check.officeOnly && !(element.ctor.prototype instanceof OfficeElement)) continue;
				const reason = await check.failure(element, { host, resources }).catch(String);
				if (reason) failing[element.tag] = reason;
			}
			const known = new Set(KNOWN_GAPS[check.id]);
			const added = Object.keys(failing).filter((tag) => !known.has(tag));
			const fixed = [...known].filter((tag) => !(tag in failing));
			expect(
				Object.fromEntries(added.map((tag) => [tag, failing[tag]])),
				`New ${check.id} failures: fix the element (a gap is never added to known-gaps.ts)`,
			).toEqual({});
			expect(fixed, `Now passing ${check.id}: remove these from known-gaps.ts`).toEqual([]);
		});
});
