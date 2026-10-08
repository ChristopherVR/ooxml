import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import {
	EMPTY_SELECTION,
	hasVisibleShapeContent,
	snapshotSelection,
	visioSelectionIsOnPage,
} from './selection';

it('keeps omitted current-page identities distinct from background-page selections', () => {
	expect(visioSelectionIsOnPage({}, '1')).toBe(true);
	expect(visioSelectionIsOnPage({ pageId: '1' }, '1')).toBe(true);
	expect(visioSelectionIsOnPage({ pageId: '2' }, '1')).toBe(false);
	expect(visioSelectionIsOnPage(null, '1')).toBe(false);
});

it('normalizes displayed identities without retaining mutable host records', () => {
	const model = structuredClone(demoDocument);
	const page = model.pages[0]!;
	const parent = page.shapes[0]!;
	parent.kind = 'group';
	parent.children = [page.shapes.splice(1, 1)[0]!];
	const child = parent.children[0]!;
	const input = [
		{ id: child.id, name: child.name, pageId: page.id },
		{ id: parent.id, name: parent.name },
	];
	const result = snapshotSelection(model, 0, input, new WeakMap());
	expect(result).toEqual([{ id: parent.id, name: parent.name }]);
	expect(Object.isFrozen(result)).toBe(true);
	expect(Object.isFrozen(result[0])).toBe(true);
	input[1]!.name = 'Changed';
	expect(result[0]!.name).toBe(parent.name);
	expect(snapshotSelection(model, 0, [], new WeakMap())).toBe(EMPTY_SELECTION);
});

it('filters hidden and non-rendering targets with the same eligibility used by painting', () => {
	const model = structuredClone(demoDocument);
	const page = model.pages[0]!;
	const shape = page.shapes[0]!;
	const visible = new WeakMap([[shape, false]]);
	expect(hasVisibleShapeContent(shape, visible)).toBe(false);
	expect(snapshotSelection(model, 0, [{ id: shape.id, name: shape.name }], visible)).toBe(
		EMPTY_SELECTION,
	);
	shape.kind = 'group';
	shape.groupDisplayMode = 0;
	expect(hasVisibleShapeContent(shape, new WeakMap())).toBe(false);
	shape.children = [page.shapes[1]!];
	expect(hasVisibleShapeContent(shape, new WeakMap())).toBe(true);
});

it('bounds supplied selections and ignores identities outside displayed pages', () => {
	expect(() =>
		snapshotSelection(demoDocument, 0, [{ id: 'x'.repeat(1025), name: 'x' }], new WeakMap()),
	).toThrow('Selection requires');
	expect(
		snapshotSelection(demoDocument, 0, [{ id: 's1', name: 'x', pageId: 'missing' }], new WeakMap()),
	).toBe(EMPTY_SELECTION);
});

it('reads host scalar getters once and does not invoke a supplied array iterator', () => {
	let reads = 0;
	const input = [
		{
			id: 's1',
			get name() {
				return ++reads === 1 ? 'Original' : 'Changed';
			},
		},
	];
	input[Symbol.iterator] = () => {
		throw new Error('Host iterator');
	};
	const selected = snapshotSelection(demoDocument, 0, input, new WeakMap());
	expect(reads).toBe(1);
	expect(selected[0]!.name).toBe('Original');
});
