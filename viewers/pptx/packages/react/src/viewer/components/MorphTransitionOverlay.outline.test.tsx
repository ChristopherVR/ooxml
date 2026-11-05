import { buildMorphTransitionPlan } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { MorphTransitionOverlay } from './MorphTransitionOverlay';

/**
 * The overlay names its halves by wrapper attributes rather than by
 * `data-element-id`, so a ghost's re-pathed outline (see `morph-shape-carry` in
 * ooxml-ui) needs rules keyed on those wrappers; without them the ghost of a
 * re-handled arrow kept its old outline over the new one.
 */
function card(id: string, adj: number): PptxElement {
	return {
		id,
		name: 'Card',
		type: 'shape',
		shapeType: 'roundRect',
		x: 100,
		y: 100,
		width: 400,
		height: 200,
		shapeAdjustments: { adj },
		shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
	} as unknown as PptxElement;
}

const slide = (id: string, elements: PptxElement[]) =>
	({ id, slideNumber: 1, elements }) as unknown as PptxSlide;

describe('morphTransitionOverlay outline rules', () => {
	it('keys the re-pathed outline of each half on its overlay wrapper', () => {
		const outgoing = slide('a', [card('a-1', 5000)]);
		const incoming = slide('b', [card('b-1', 50000)]);
		const plan = buildMorphTransitionPlan(outgoing, incoming, 2000)!;
		const html = renderToStaticMarkup(
			<MorphTransitionOverlay
				canvasSize={{ width: 1280, height: 720 }}
				scale={1}
				morphPlan={plan}
				outgoingSlide={outgoing}
				incomingSlide={incoming}
				containerRef={{ current: null }}
			/>,
		);
		expect(html).toMatch(
			/\[data-pptx-morph-outgoing=(?:"|&quot;)a-1(?:"|&quot;)\] \[data-pptx-outline\] path \{ animation: pptx-morph-carry-g0-a1-outline /u,
		);
		expect(html).toContain('<svg');
	});
});
