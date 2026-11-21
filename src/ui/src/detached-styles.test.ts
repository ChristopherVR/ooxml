import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { releaseDetachedStyles } from './detached-styles';

const nextTask = () => new Promise((resolve) => setTimeout(resolve, 0));
/** The text of every stylesheet the document adopted, so a test sees each change. */
const styles = () =>
	document.adoptedStyleSheets.map((sheet) =>
		Array.from(sheet.cssRules, (rule) => rule.cssText).join(''),
	);

describe('releaseDetachedStyles', () => {
	let before: CSSStyleSheet[];
	beforeEach(() => {
		before = [...document.adoptedStyleSheets];
	});
	afterEach(() => {
		document.adoptedStyleSheets = before;
		document.body.replaceChildren();
	});

	it('changes the document styles once an element stays removed past the task', async () => {
		const element = document.body.appendChild(document.createElement('div'));
		const start = styles();
		element.remove();
		releaseDetachedStyles(element);
		expect(styles()).toEqual(start);
		await nextTask();
		expect(styles()).not.toEqual(start);
		const once = styles();
		releaseDetachedStyles(element);
		await nextTask();
		expect(styles()).not.toEqual(once);
	});

	it('leaves the styles alone for an element moved within the task', async () => {
		const element = document.body.appendChild(document.createElement('div'));
		const start = styles();
		element.remove();
		releaseDetachedStyles(element);
		document.body.append(element);
		await nextTask();
		expect(styles()).toEqual(start);
	});

	it('makes one change for every element removed in the same task', async () => {
		// Two removals: two changes would toggle the rule back out.
		const elements = [1, 2].map(() => document.body.appendChild(document.createElement('div')));
		// The shared worker may already have adopted the sheet alongside unrelated styles.
		// Clear the list so the next release exposes the exact sheet this test measures.
		document.adoptedStyleSheets = [];
		releaseDetachedStyles(document.createElement('div'));
		await nextTask();
		const sheet = document.adoptedStyleSheets.at(-1)!;
		const rules = sheet.cssRules.length;
		document.adoptedStyleSheets = [];
		for (const element of elements) {
			element.remove();
			releaseDetachedStyles(element);
		}
		await nextTask();
		expect(document.adoptedStyleSheets).toEqual([sheet]);
		expect(sheet.cssRules.length).toBe(1 - rules);
	});

	it('adopts its stylesheet again after the host replaced the list', async () => {
		const element = document.createElement('div');
		releaseDetachedStyles(element);
		await nextTask();
		document.adoptedStyleSheets = [];
		releaseDetachedStyles(element);
		await nextTask();
		expect(document.adoptedStyleSheets).toHaveLength(1);
	});
});
