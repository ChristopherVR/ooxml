import { EditorState } from 'prosemirror-state';
import { expect, it } from 'vitest';
import { expandStack } from './arrange-group';
import { parsePlacement } from './print-float-locate';
import { schema } from './schema';

it('reserves child heights by shifting only drawings above the group in its text layer', () => {
	const entries = [
		{ height: 400 },
		{ height: 100 },
		{ height: 200 },
		{ height: 200, group: true },
		{ height: 200 },
		{ height: 300, behindText: true },
		{ height: 300, anchored: false },
	];
	const images = entries.map(({ height, behindText, anchored = true }) =>
		schema.node('image', {
			anchored,
			placement: JSON.stringify({ relativeHeight: height, behindText }),
		}),
	);
	const doc = schema.node('doc', null, [schema.node('paragraph', null, images)]);
	const tr = EditorState.create({ doc }).tr;
	const first = expandStack(
		tr,
		1 + entries.findIndex((entry) => entry.group),
		{ wrap: 'square', relativeHeight: 200 },
		3,
	);
	expect(first).toBe(200);
	const heights: Array<number | undefined> = [];
	tr.doc.firstChild?.forEach((node) =>
		heights.push(parsePlacement(node.attrs.placement)?.relativeHeight),
	);
	expect(heights).toEqual([402, 100, 200, 200, 202, 300, 300]);
});
