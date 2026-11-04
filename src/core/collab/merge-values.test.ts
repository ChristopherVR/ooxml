import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { mergeSnapshotMaps, mergeSnapshotValue } from './merge-values';

describe('saved snapshot reconciliation', () => {
	it('combines independent fields and preserves deletions', () => {
		const issues: string[] = [];
		expect(
			mergeSnapshotValue(
				{ value: 1, style: 'plain', removed: true },
				{ value: 2, style: 'plain' },
				{ value: 1, style: 'bold', removed: true },
				'cell',
				issues,
			),
		).toEqual({ value: 2, style: 'bold' });
		expect(issues).toEqual([]);
	});
	it('does not invent ordering or overwrite overlapping values', () => {
		for (const [base, local, remote] of [
			[1, 2, 3],
			[[1], [1, 2], [0, 1]],
			[{ value: 1 }, undefined, { value: 2 }],
		]) {
			const issues: string[] = [];
			mergeSnapshotValue(base, local, remote, 'field', issues);
			expect(issues).toEqual(['Concurrent changes overlap at field.']);
		}
	});
	it('reconciles atomic cell JSON using Yjs branches seeded from one baseline', () => {
		const base = new Y.Doc();
		base.getMap('cells').set('A1', { value: 1, style: 'plain' });
		const local = new Y.Doc(),
			remote = new Y.Doc();
		for (const doc of [local, remote]) Y.applyUpdate(doc, Y.encodeStateAsUpdate(base));
		local.getMap('cells').set('A1', { value: 2, style: 'plain' });
		remote.getMap('cells').set('A1', { value: 1, style: 'bold' });
		const merged = mergeSnapshotMaps(base, local, remote, ['cells']);
		expect(merged.issues).toEqual([]);
		expect(merged.doc.getMap('cells').get('A1')).toEqual({ value: 2, style: 'bold' });
		for (const doc of [base, local, remote, merged.doc]) doc.destroy();
	});
});
