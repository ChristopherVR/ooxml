import type { PptxSmartArtData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { applySmartArtTextPaneRequest, smartArtTextPaneView } from './smartart-text-pane-model';

const t = (key: string, params?: Record<string, string | number>) =>
	params ? `${key}${JSON.stringify(params)}` : key;

/** A parsed diagram: top-level points hang off the document point, which is not a node. */
const data = (): PptxSmartArtData => ({
	resolvedLayoutType: 'list',
	nodes: [
		{ id: 'a', text: 'Alpha', parentId: 'doc' },
		{ id: 'a1', text: 'Child', parentId: 'a', style: { bold: true } },
		{ id: 'b', text: '', parentId: 'doc' },
	],
	connections: [
		{ sourceId: 'doc', destId: 'a', type: 'parOf', srcOrd: 0 },
		{ sourceId: 'a', destId: 'a1', type: 'parOf', srcOrd: 0 },
		{ sourceId: 'doc', destId: 'b', type: 'parOf', srcOrd: 1 },
		{ sourceId: 'a', destId: 'b', type: 'sibTrans' },
	],
});

describe('smartArtTextPaneView', () => {
	it('numbers top-level lines, bullets deeper ones and sets the tool flags', () => {
		const view = smartArtTextPaneView(data(), t);
		expect(view.lines.map((l) => [l.id, l.level, l.marker])).toEqual([
			['a', 0, '1'],
			['a1', 1, '\u2022'],
			['b', 0, '2'],
		]);
		expect(view.lines[0]!.label).toBe('pptx.smartArt.itemLabel{"index":1,"text":"Alpha"}');
		expect(view.lines[2]!.label).toContain('pptx.smartArt.emptyNode');
		expect(view.lines[1]!.style).toEqual({ bold: true });
		expect(view.lines.map((l) => [l.canMoveUp, l.canMoveDown, l.canAddChild])).toEqual([
			[false, true, true],
			[false, false, false],
			[true, false, true],
		]);
		expect(view.notes[0]).toContain('pptx.smartart.extraConnections{"count":1} (sibTrans)');
		expect(view.notes.at(-1)).toBe('pptx.smartart.tabHint');
		expect(view.appendDisabled).toBe(false);
	});

	it('refuses removing a shape the layout needs', () => {
		const pyramid = { ...data(), resolvedLayoutType: 'pyramid' as const };
		const view = smartArtTextPaneView(pyramid, t);
		expect(view.lines.filter((l) => l.level === 0).every((l) => !l.canRemove)).toBe(true);
		expect(
			applySmartArtTextPaneRequest(pyramid, { kind: 'remove', id: 'b' }, 'Sub'),
		).toBeUndefined();
	});
});

describe('applySmartArtTextPaneRequest', () => {
	it('edits text and skips an unchanged one', () => {
		const edit = applySmartArtTextPaneRequest(data(), { kind: 'text', id: 'b', text: 'Beta' }, 'S');
		expect(edit?.data.nodes[2]!.text).toBe('Beta');
		expect(
			applySmartArtTextPaneRequest(data(), { kind: 'text', id: 'a', text: 'Alpha' }, 'S'),
		).toBeUndefined();
	});

	it('adds after, removes, moves and changes level with the line to focus', () => {
		const added = applySmartArtTextPaneRequest(data(), { kind: 'add', id: 'a' }, 'S')!;
		expect(added.data.nodes).toHaveLength(4);
		expect(added.focusNodeId).toBe(added.data.nodes[1]!.id);
		const removed = applySmartArtTextPaneRequest(data(), { kind: 'remove', id: 'b' }, 'S')!;
		expect(removed.data.nodes.map((n) => n.id)).toEqual(['a', 'a1']);
		const moved = applySmartArtTextPaneRequest(
			data(),
			{ kind: 'move', id: 'b', direction: -1 },
			'S',
		)!;
		expect(moved.focusNodeId).toBe('b');
		expect(moved.data.nodes.map((n) => n.id)).toEqual(['b', 'a1', 'a']);
		const demoted = applySmartArtTextPaneRequest(data(), { kind: 'demote', id: 'b' }, 'S')!;
		expect(demoted.data.nodes.find((n) => n.id === 'b')?.parentId).toBe('a');
		const promoted = applySmartArtTextPaneRequest(data(), { kind: 'promote', id: 'a1' }, 'S')!;
		expect(promoted.data.nodes.find((n) => n.id === 'a1')?.parentId).toBe('doc');
		expect(applySmartArtTextPaneRequest(data(), { kind: 'promote', id: 'a' }, 'S')).toBeUndefined();
	});

	it('adds a sub-item and appends a top-level item, focusing the new node', () => {
		const child = applySmartArtTextPaneRequest(data(), { kind: 'addChild', id: 'b' }, 'Sub')!;
		const node = child.data.nodes.find((n) => n.id === child.focusNodeId);
		expect(node).toMatchObject({ parentId: 'b', text: 'Sub' });
		const appended = applySmartArtTextPaneRequest(data(), { kind: 'append' }, 'Sub')!;
		expect(appended.data.nodes.find((n) => n.id === appended.focusNodeId)?.parentId).toBe('doc');
	});

	it('sets a node style', () => {
		const edit = applySmartArtTextPaneRequest(
			data(),
			{ kind: 'style', id: 'a', style: { fillColor: '#ff0000' } },
			'S',
		)!;
		expect(edit.data.nodes[0]!.style?.fillColor).toBe('#ff0000');
	});
});
