import { describe, expect, it } from 'vitest';

import { parseOrderedXml } from '../engine/ordered-xml';
import type { OrderedXmlElement } from '../engine/ordered-xml';
import type { DiagramChoose, DiagramLayoutNode } from '../model';
import { groupedChildren, orderedXmlView } from './smartart-choose-xml';
import { parseOrderedWhen } from './smartart-layout-control-flow';
import { chooseAlgorithm, chooseAlgType } from './smartart-layout-interpreter-choose-algorithm';
import { structuralChooseAlgDepth } from './smartart-layout-interpreter-choose-depth';

/** A layout node whose `dgm:choose` children are parsed from `xml` with the ordered reader. */
function nodeFrom(xml: string): DiagramLayoutNode<OrderedXmlElement> {
	const root = parseOrderedXml(xml)!;
	const choose = root.children
		.filter((child) => child.name === 'choose')
		.map((entry): DiagramChoose<OrderedXmlElement> => {
			const otherwise = entry.children.find((child) => child.name === 'else');
			return {
				when: entry.children
					.filter((child) => child.name === 'if')
					.map(parseOrderedWhen)
					.filter((value) => value !== undefined),
				...(otherwise ? { otherwise: { rawXml: otherwise } } : {}),
				rawXml: entry,
			};
		});
	return { name: root.attrs['name'] ?? '', choose };
}

const NS = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';

describe('ordered-XML choose walkers', () => {
	it('groups same-named children in first-appearance order', () => {
		const root = parseOrderedXml('<a><x n="1"/><y/><x n="2"/></a>')!;
		expect(groupedChildren(root).map((child) => child.attrs['n'] ?? child.name)).toEqual([
			'1',
			'2',
			'y',
		]);
	});

	it('resolves the winning branch algorithm and its params', () => {
		const node = nodeFrom(`<dgm:layoutNode ${NS} name="root"><dgm:choose>
			<dgm:if func="cnt" op="lte" val="2"><dgm:alg type="cycle"/></dgm:if>
			<dgm:else><dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg></dgm:else>
		</dgm:choose></dgm:layoutNode>`);
		expect(chooseAlgType(orderedXmlView, node, 2)).toBe('cycle');
		expect(chooseAlgorithm(orderedXmlView, node, 5)).toEqual({
			type: 'lin',
			parameters: [{ type: 'linDir', value: 'fromT' }],
		});
	});

	it('evaluates a nested choose instead of taking its first branch', () => {
		const node = nodeFrom(`<dgm:layoutNode ${NS}><dgm:choose><dgm:else><dgm:choose>
			<dgm:if func="cnt" op="equ" val="1"><dgm:alg type="cycle"/></dgm:if>
			<dgm:else><dgm:alg type="snake"/></dgm:else>
		</dgm:choose></dgm:else></dgm:choose></dgm:layoutNode>`);
		expect(chooseAlgType(orderedXmlView, node, 1)).toBe('cycle');
		expect(chooseAlgType(orderedXmlView, node, 3)).toBe('snake');
	});

	it('counts the layout nodes crossed before the structural algorithm', () => {
		const node = nodeFrom(`<dgm:layoutNode ${NS}><dgm:choose><dgm:else>
			<dgm:layoutNode name="outer"><dgm:layoutNode name="inner"><dgm:alg type="lin"/>
			</dgm:layoutNode></dgm:layoutNode>
		</dgm:else></dgm:choose></dgm:layoutNode>`);
		expect(structuralChooseAlgDepth(orderedXmlView, node, 3, {})).toEqual({
			depth: 2,
			crossedNames: ['outer', 'inner'],
		});
	});
});
