// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { assignKeyTips, runKeyTips, type KeyTipTarget } from './keytip-run';

const target = (name: string, key = '', activate = () => {}): KeyTipTarget => {
	const element = document.createElement('button');
	element.setAttribute('aria-label', name);
	document.body.append(element);
	return { key, element, activate };
};
const nameOf = (item: KeyTipTarget) => item.element.getAttribute('aria-label') ?? '';

afterEach(() => document.body.replaceChildren());

describe('assignKeyTips', () => {
	it('keeps the product keys, re-derives a clash and never makes one tip the prefix of another', () => {
		const items = [target('Bold', '1'), target('Italic', '1'), target('Font', 'F'), target('Fill')];
		assignKeyTips(items, nameOf);
		const keys = items.map((item) => item.key);
		expect(keys[0]).toBe('1');
		expect(keys[2]).toBe('F');
		expect(new Set(keys).size).toBe(keys.length);
		for (const key of keys)
			expect(keys.filter((other) => other !== key && other.startsWith(key))).toEqual([]);
	});

	it('falls back to numbers for a label with nothing to derive from', () => {
		const items = [target('!'), target('?')];
		assignKeyTips(items, nameOf);
		expect(items.map((item) => item.key)).toEqual(['ZZ', '10']);
	});
});

describe('runKeyTips', () => {
	it('shows a badge per target, narrows as letters are typed and runs the exact match', () => {
		const run = vi.fn();
		const done = vi.fn();
		const items = [target('Bold', 'B1', run), target('Copy', 'C')];
		runKeyTips(document.body, items, done, 'tip');
		const badges = () => [...document.querySelectorAll<HTMLElement>('.tip')];
		expect(badges().map((badge) => badge.textContent)).toEqual(['B1', 'C']);
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'b', bubbles: true }));
		expect(badges().map((badge) => badge.hidden)).toEqual([false, true]);
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: '1', bubbles: true }));
		expect(run).toHaveBeenCalledOnce();
		expect(done).toHaveBeenCalledOnce();
		expect(badges()).toEqual([]);
	});

	it('ends on Escape and on a key with no match', () => {
		const done = vi.fn();
		runKeyTips(document.body, [target('Bold', 'B')], done, 'tip');
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(done).toHaveBeenCalledOnce();
		runKeyTips(document.body, [target('Bold', 'B')], done, 'tip');
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'q', bubbles: true }));
		expect(done).toHaveBeenCalledTimes(2);
	});
});
