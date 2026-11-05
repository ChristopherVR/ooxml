import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import { buildMorphTransitionPlan } from './morph-plan';
import { morphCopyId } from './morph-text-body';
import type { MorphTextMeasure } from './morph-text-frame';

/** Every character advances by a per-face fraction of the font size. */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0', family = ''] = /([\d.]+)px\s+"?([^",]+)/u.exec(font) ?? [];
	return [...text].length * Number(size) * (/Arial/u.test(family) ? 0.52 : 0.47);
};

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

/** A rounded rectangle with a fill, an outline and one centred line. */
function card(
	id: string,
	style: { fontFamily: string; fontSize: number },
	box: Partial<PptxElement>,
	shapeStyle: Record<string, unknown> = {},
): PptxElement {
	return {
		id,
		type: 'shape',
		name: '!!Card',
		shapeType: 'rect',
		x: 107,
		y: 133,
		width: 400,
		height: 160,
		text: 'Lorem ipsum',
		textSegments: [{ text: 'Lorem ipsum', style: { ...style, color: '#FFFFFF' } }],
		textStyle: { ...style, align: 'center', vAlign: 'middle', textWrap: 'none' },
		shapeStyle: {
			fillMode: 'solid',
			fillColor: '#2070C0',
			strokeColor: '#FFFFFF',
			strokeWidth: 4,
			...shapeStyle,
		},
		...box,
	} as PptxElement;
}

// PowerPoint 16 (`CreateVideo`, 720p): a rounded rectangle moving and growing
// from 300 x 120 pt to 420 x 180 pt while its text changes Calibri 28 pt ->
// Arial 40 pt. The box follows the morph curve; the text stays centred on it,
// grows on the same curve and dissolves glyph over glyph at full strength.
describe('a shape whose text changes while its box moves', () => {
	const from = card('a', { fontFamily: 'Calibri', fontSize: 37.33 }, {});
	const to = card(
		'b',
		{ fontFamily: 'Arial', fontSize: 53.33 },
		{ x: 613, y: 400, width: 560, height: 240 },
	);
	const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'object', {
		measureText: measure,
	})!;

	it('paints the body and the text as two crossfade groups of copies', () => {
		expect(plan.crossfadeGroups.map((group) => [group.outgoing.id, group.incoming.id])).toEqual([
			[morphCopyId('a', 'body'), morphCopyId('b', 'body')],
			[morphCopyId('a', 'text'), morphCopyId('b', 'text')],
		]);
		const [body, text] = plan.crossfadeGroups as [MorphCrossfadeGroup, MorphCrossfadeGroup];
		expect((body.incoming as { text?: string }).text).toBe('');
		const textStyle = (text.incoming as { shapeStyle?: Record<string, unknown> }).shapeStyle;
		expect(textStyle).toMatchObject({ fillMode: 'none', strokeWidth: 0 });
		expect(textStyle?.fillColor).toBeUndefined();
		expect(textStyle?.strokeColor).toBeUndefined();
		expect(plan.morphCopyParts?.get(morphCopyId('b', 'text'))).toBe('text');
		expect(plan.textPieceSources?.get(morphCopyId('b', 'body'))).toBe('b');
	});

	it('hides the pair itself and stacks the text over the body', () => {
		expect(plan.outgoingElements).toHaveLength(0);
		expect(plan.incomingAnimations.get('b')).toContain('pptx-morph-lifted-hidden');
		expect(plan.overlayStack).toEqual([
			{ kind: 'crossfade', index: 0 },
			{ kind: 'crossfade', index: 1 },
		]);
	});

	it('interpolates the body box and carries the text on its own mapping', () => {
		const css = plan.keyframesCss;
		const body = css.split(`@keyframes pptx-morph-body-${'b'}morphbody`)[1] ?? '';
		expect(body).toContain(`scale(0.714, 0.667)`);
		const text = css.split('@keyframes pptx-morph-text-bmorphtext')[1] ?? '';
		const ratio = (53.33 * 0.52) / (37.33 * 0.47);
		expect(text).toContain(`scale(${Math.round((1 / ratio) * 1000) / 1000}, `);
		// Both centred on their boxes, so the text's start translation is the box's.
		expect(text).toContain('translate(-586px, -307px)');
	});

	it('leaves a box that only moves to the plain interpolation', () => {
		const moved = card('c', { fontFamily: 'Calibri', fontSize: 37.33 }, { x: 613, y: 400 });
		const only = buildMorphTransitionPlan(slide([from]), slide([moved]), 2000, 'object', {
			measureText: measure,
		})!;
		expect(only.morphCopyParts).toBeUndefined();
	});

	it('refuses a shape with an effect it could not paint once', () => {
		const glowing = (element: PptxElement) =>
			({
				...element,
				shapeStyle: { ...(element as { shapeStyle?: object }).shapeStyle, glowColor: '#fff' },
			}) as PptxElement;
		const refused = buildMorphTransitionPlan(
			slide([glowing(from)]),
			slide([glowing(to)]),
			2000,
			'object',
			{ measureText: measure },
		)!;
		expect(refused.morphCopyParts).toBeUndefined();
	});

	it('leaves word and character morphs alone', () => {
		const word = buildMorphTransitionPlan(slide([from]), slide([to]), 2000, 'word', {
			measureText: measure,
		})!;
		expect(word.morphCopyParts).toBeUndefined();
	});
});

// PowerPoint 16 (`CreateVideo`, 720p), the same card while its body does more:
// a shadow stays under the body, rounder corners round the body only, and a
// flip turns the body edge-on while the text travels upright, glyph over glyph.
describe('a shape whose body does more than resize', () => {
	const from = card('a', { fontFamily: 'Calibri', fontSize: 37.33 }, {});
	const target = { x: 613, y: 400, width: 560, height: 240 };
	const to = card('b', { fontFamily: 'Arial', fontSize: 53.33 }, target);
	const plan = (a: PptxElement, b: PptxElement) =>
		buildMorphTransitionPlan(slide([a]), slide([b]), 2000, 'object', { measureText: measure })!;
	const block = (css: string, name: string) =>
		css.split(`@keyframes ${name}`)[1]?.split('@keyframes')[0] ?? '';

	it('leaves an outer shadow under an opaque fill to the body', () => {
		const shadow = { shadowColor: '#000000', shadowBlur: 8, shadowOffsetX: 10 };
		const split = plan(
			card('a', { fontFamily: 'Calibri', fontSize: 37.33 }, {}, shadow),
			card('b', { fontFamily: 'Arial', fontSize: 53.33 }, target, shadow),
		);
		const [body, text] = split.crossfadeGroups as [MorphCrossfadeGroup, MorphCrossfadeGroup];
		const styleOf = (element: PptxElement) =>
			(element as { shapeStyle?: Record<string, unknown> }).shapeStyle;
		expect(styleOf(body.incoming)?.shadowColor).toBe('#000000');
		expect(styleOf(text.incoming)?.shadowColor).toBeUndefined();
		// Over a see-through fill the text would cast it too: kept whole.
		const faint = { ...shadow, fillOpacity: 0.5 };
		const whole = plan(
			card('a', { fontFamily: 'Calibri', fontSize: 37.33 }, {}, faint),
			card('b', { fontFamily: 'Arial', fontSize: 53.33 }, target, faint),
		);
		expect(whole.morphCopyParts).toBeUndefined();
	});

	it('tweens the outline of a body whose adjustments change', () => {
		const rounded = (element: PptxElement, adj: number) =>
			({ ...element, shapeType: 'roundRect', shapeAdjustments: { adj } }) as PptxElement;
		const split = plan(rounded(from, 10000), rounded(to, 40000));
		expect(split.morphCopyParts?.get(morphCopyId('b', 'text'))).toBe('text');
		const animation = split.overlayIncomingAnimations.get(morphCopyId('b', 'body')) ?? '';
		expect(animation).toContain('pptx-morph-geo-');
		expect(split.overlayIncomingAnimations.get(morphCopyId('b', 'text'))).not.toContain('geo');
	});

	it('turns the body over in 3D and keeps the text upright', () => {
		for (const flip of ['flipHorizontal', 'flipVertical'] as const) {
			const split = plan(from, { ...to, [flip]: true } as PptxElement);
			const css = split.keyframesCss;
			const body = block(css, 'pptx-morph-body-bmorphbody');
			const turn = flip === 'flipHorizontal' ? 'rotateY' : 'rotateX';
			expect(body).toMatch(/from \{[^}]*perspective\([^}]*\}\s*to \{[^}]*perspective/u);
			expect(body).toContain(`${turn}(0deg)`);
			expect(body).toMatch(new RegExp(`${turn}[(]-?180deg`, 'u'));
			expect(block(css, 'pptx-morph-body-amorphbody')).toContain('perspective(');
			expect(block(css, 'pptx-morph-text-bmorphtext')).not.toContain('perspective');
			expect(block(css, 'pptx-morph-text-bmorphtext')).not.toContain('scaleX');
			const text = (split.crossfadeGroups[1] as MorphCrossfadeGroup).incoming;
			expect(text.flipHorizontal).toBe(false);
		}
	});

	it('does not turn the body of a pair whose flips are unchanged', () => {
		const split = plan(from, to);
		expect(split.keyframesCss).not.toContain('perspective');
	});

	it('turns the body with the box and the text about its middle', () => {
		const split = plan(from, { ...to, rotation: 30 } as PptxElement);
		const css = split.keyframesCss;
		expect(block(css, 'pptx-morph-body-bmorphbody')).toContain('rotate(0deg)');
		expect(block(css, 'pptx-morph-body-bmorphbody')).toContain('rotate(30deg)');
		const text = block(css, 'pptx-morph-text-bmorphtext');
		expect(text.match(/\) translate\(/gu)).toHaveLength(2);
	});
});
