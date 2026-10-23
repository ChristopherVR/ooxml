/**
 * group-press-target.test.ts: selecting inside a group, the sequence every
 * binding runs on a press or a double-click. The first press selects the
 * group, a press on a member of the selected group selects that member, a
 * press on another member moves to it, the selected member stays selected,
 * and a double-click goes straight to the member under the pointer. Moved
 * from the Angular canvas tests with the functions.
 */

import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	groupDrillChain,
	resolveGroupDoubleClickTarget,
	resolveGroupPressTarget,
} from './group-press-target';

const card = (id: string, x: number) =>
	({
		id,
		type: 'shape',
		x,
		y: 0,
		width: 100,
		height: 60,
		shapeType: 'roundRect',
		text: `${id} title`,
	}) as unknown as PptxElement;
// A group at (200, 100) with two cards side by side (children are group-relative).
const cards = {
	id: 'cards',
	type: 'group',
	x: 200,
	y: 100,
	width: 220,
	height: 60,
	children: [card('a', 0), card('b', 120)],
} as unknown as PptxElement;
const plain = card('plain', 0);
const slideElements: PptxElement[] = [plain, cards];

// Slide-space points over card a, card b, and the gap between them.
const overA = { x: 250, y: 130 };
const overB = { x: 370, y: 130 };
const overGap = { x: 310, y: 130 };

describe('resolveGroupPressTarget', () => {
	it('selects the group on the first press', () => {
		expect(resolveGroupPressTarget(slideElements, 'cards', overA, [])).toBe('cards');
		expect(resolveGroupPressTarget(slideElements, 'cards', overA, ['plain'])).toBe('cards');
	});

	it('selects the member under the pointer once the group is selected', () => {
		expect(resolveGroupPressTarget(slideElements, 'cards', overA, ['cards'])).toBe('a');
		expect(resolveGroupPressTarget(slideElements, 'cards', overB, ['cards'])).toBe('b');
	});

	it('moves to another member of the entered group', () => {
		expect(resolveGroupPressTarget(slideElements, 'cards', overB, ['a'])).toBe('b');
	});

	it('keeps the selected member selected (its text opens on release)', () => {
		expect(resolveGroupPressTarget(slideElements, 'cards', overA, ['a'])).toBe('a');
	});

	it('keeps the group on a press between its members', () => {
		expect(resolveGroupPressTarget(slideElements, 'cards', overGap, ['cards'])).toBe('cards');
	});

	it('leaves non-groups, rotated groups and multi-selections alone', () => {
		expect(resolveGroupPressTarget(slideElements, 'plain', { x: 10, y: 10 }, ['plain'])).toBe(
			'plain',
		);
		const rotated = [{ ...cards, rotation: 30 } as PptxElement];
		expect(resolveGroupPressTarget(rotated, 'cards', overA, ['cards'])).toBe('cards');
		expect(groupDrillChain(rotated, 'cards', overA)).toBeNull();
		expect(resolveGroupPressTarget(slideElements, 'cards', overA, ['cards', 'plain'])).toBe(
			'cards',
		);
	});
});

describe('locked group members', () => {
	const withLocks = (first: Record<string, unknown>, group: Record<string, unknown> = {}) => {
		const g = {
			...(cards as unknown as Record<string, unknown>),
			...group,
			children: [{ ...card('a', 0), ...first }, card('b', 120)],
		};
		return [plain, g as unknown as PptxElement];
	};

	it('never drills into a noSelect member (press or double-click)', () => {
		const elements = withLocks({ locks: { noSelect: true } });
		expect(resolveGroupPressTarget(elements, 'cards', overA, ['cards'])).toBe('cards');
		expect(resolveGroupDoubleClickTarget(elements, 'cards', overA)).toBe('cards');
		// The unlocked sibling still resolves.
		expect(resolveGroupDoubleClickTarget(elements, 'cards', overB)).toBe('b');
	});

	it('never drills into a noDrilldown group', () => {
		const elements = withLocks({}, { locks: { noDrilldown: true } });
		expect(resolveGroupPressTarget(elements, 'cards', overA, ['cards'])).toBe('cards');
		expect(resolveGroupDoubleClickTarget(elements, 'cards', overA)).toBe('cards');
	});
});

describe('resolveGroupDoubleClickTarget', () => {
	it('goes straight to the member under the pointer', () => {
		expect(resolveGroupDoubleClickTarget(slideElements, 'cards', overB)).toBe('b');
	});

	it('stays on the group between members, and on a plain shape', () => {
		expect(resolveGroupDoubleClickTarget(slideElements, 'cards', overGap)).toBe('cards');
		expect(resolveGroupDoubleClickTarget(slideElements, 'plain', { x: 1, y: 1 })).toBe('plain');
	});
});
