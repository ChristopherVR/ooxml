import { describe, expect, it } from 'vitest';
import type { DataPoint } from './data-points';
import { buildPresentationTree, flattenPresTree } from './pres-tree';
import type { LdDefinition, LdIterator, LdLayoutNode, LdStatement } from './layout-def-types';

const iterator = (axis: string): LdIterator => ({
	axis: [axis],
	ptType: ['node'],
	hideLastTrans: [false],
	st: [1],
	cnt: [0],
	step: [1],
});

const point = (id: string, children: DataPoint[] = []): DataPoint => ({
	id,
	type: id === 'doc' ? 'doc' : 'node',
	children,
});

function definition(body: LdStatement[]): LdDefinition {
	const root: LdLayoutNode = { kind: 'layoutNode', name: 'root', chOrder: 't', body };
	return { root, forEachByName: new Map() };
}

describe('presentation tree budgets', () => {
	it('stops expanding once the total node budget is reached', () => {
		// Two nested `ch` loops of 400 points each would create 160,000 nodes.
		const leaf = (): DataPoint =>
			point(
				'group',
				Array.from({ length: 400 }, (_, i) => point(`n${i}`)),
			);
		const doc = point(
			'doc',
			Array.from({ length: 400 }, () => leaf()),
		);
		const inner: LdStatement = {
			kind: 'forEach',
			iterator: iterator('ch'),
			body: [{ kind: 'layoutNode', name: 'item', chOrder: 't', body: [] }],
		};
		const outer: LdStatement = { kind: 'forEach', iterator: iterator('ch'), body: [inner] };
		const tree = buildPresentationTree(definition([outer]), doc);
		expect(tree).toBeDefined();
		expect(flattenPresTree(tree!)).toHaveLength(100_000);
	});

	it('keeps variable maps free of inherited names', () => {
		// JSON.parse yields an own `__proto__` key, the shape a hostile varLst can produce.
		const vars = JSON.parse('{"__proto__":"hostile"}') as Record<string, string>;
		const doc = point('doc', [point('a')]);
		const tree = buildPresentationTree(definition([{ kind: 'varLst', vars }]), doc);
		expect(Object.getPrototypeOf(tree!.vars)).toBeNull();
		expect(tree!.vars['constructor']).toBeUndefined();
		expect(Object.keys(tree!.vars)).toContain('__proto__');
	});
});
