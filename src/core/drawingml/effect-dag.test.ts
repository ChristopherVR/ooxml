import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/index';
import { interpretDrawingEffectDag, readDrawingEffectContainer } from './effect-dag';
import { DOM_DRAWING_VIEW } from './node-view';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const dag = (body: string, attributes = '') =>
	readDrawingEffectContainer(
		DOM_DRAWING_VIEW,
		parseXml(`<a:effectDag xmlns:a="${A}" ${attributes}>${body}</a:effectDag>`).documentElement,
	)!;

describe('readDrawingEffectContainer', () => {
	it('types containers, list effects and colour operations with schema defaults', () => {
		const graph = dag(
			`<a:cont type="tree" name="inner"><a:grayscl/><a:outerShdw blurRad="38100"><a:srgbClr val="000000"/></a:outerShdw></a:cont>
			<a:alphaModFix/><a:lum bright="20000"/><a:hsl hue="5400000" sat="-50000"/><a:xfrm sx="50000" tx="12700"/>
			<a:clrChange useA="0"><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom><a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>
			<a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"/></a:duotone><a:effect ref="fillLine"/><a:foo/>`,
			'type="sib"',
		);
		expect(graph.type).toBe('sib');
		const [inner, alpha, lum, hsl, xfrm, change, duotone, reference, unknown] = graph.children;
		expect(inner).toMatchObject({ kind: 'container', type: 'tree', name: 'inner' });
		expect(inner?.kind === 'container' && inner.children.map((child) => child.kind)).toEqual([
			'colorOp',
			'effect',
		]);
		expect(alpha).toEqual({ kind: 'alphaModFix', amount: 1 });
		expect(lum).toMatchObject({ op: { kind: 'luminance', brightness: 0.2, contrast: 0 } });
		expect(hsl).toMatchObject({ op: { kind: 'hsl', hue: 90, saturation: -0.5, luminance: 0 } });
		expect(xfrm).toMatchObject({ kind: 'transform', scaleX: 0.5, scaleY: 1, offsetXEmu: 12700 });
		expect(change).toMatchObject({
			op: { kind: 'colorChange', useAlpha: false, from: { value: 'FFFFFF' } },
		});
		expect(
			duotone?.kind === 'colorOp' && duotone.op.kind === 'duotone' && duotone.op.colors,
		).toHaveLength(2);
		expect(reference).toEqual({ kind: 'reference', ref: 'fillLine' });
		expect(unknown).toEqual({ kind: 'unsupported', name: 'foo' });
	});
});

describe('interpretDrawingEffectDag', () => {
	it('collects effects and colour operations along nested containers in order', () => {
		const paint = interpretDrawingEffectDag(
			dag(
				`<a:cont type="tree"><a:glow rad="63500"><a:schemeClr val="accent1"/></a:glow>
					<a:cont><a:softEdge rad="12700"/></a:cont></a:cont><a:tint hue="1200000" amt="40000"/>`,
				'type="tree"',
			),
		);
		expect(paint.effects.map(({ effect }) => effect.kind)).toEqual(['glow', 'softEdge']);
		expect(paint.colorOps).toMatchObject([{ kind: 'tint', hue: 20, amount: 0.4 }]);
		expect(paint.opacity).toBe(1);
		expect(paint.approximations).toEqual([]);
	});

	it('multiplies alpha, reduces alphaMod to its container alpha and paints blend containers', () => {
		const paint = interpretDrawingEffectDag(
			dag(
				`<a:alphaModFix amt="50000"/><a:alphaMod><a:cont><a:alphaModFix amt="50000"/><a:glow rad="1"/></a:cont></a:alphaMod>
				<a:blend blend="mult"><a:cont><a:blur rad="25400"/></a:cont></a:blend>`,
				'type="tree"',
			),
		);
		expect(paint.opacity).toBeCloseTo(0.25);
		expect(paint.effects.map(({ effect }) => effect.kind)).toEqual(['blur']);
		expect(paint.blend).toBe('mult');
		expect(paint.approximations).toEqual([
			'alphaMod reduced to its container alpha',
			'blend mult painted as over',
		]);
	});

	it('lets alphaRepl replace the alpha and names each approximation once', () => {
		const paint = interpretDrawingEffectDag(
			dag(
				`<a:alphaModFix amt="20000"/><a:alphaRepl a="70000"/><a:relOff tx="10000" ty="-5000"/>
				<a:relOff tx="1"/><a:fill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fill><a:effect ref="x"/><a:bar/>`,
			),
		);
		expect(paint.opacity).toBeCloseTo(0.7);
		expect(paint.offset).toEqual({ tx: 0.1, ty: -0.05 });
		expect(paint.approximations).toEqual([
			'sibling outputs composited as one paint',
			'relOff reported, not applied',
			'fill effect not applied',
			'effect references not applied',
			'unsupported bar',
		]);
	});
});
