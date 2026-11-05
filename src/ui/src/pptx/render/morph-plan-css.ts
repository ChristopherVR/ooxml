/**
 * Render a morph plan as CSS: the scoped `animation` rules and the outline,
 * custom-geometry and connector-strand descendant rules a binding injects next
 * to {@link MorphTransitionPlan.keyframesCss}.
 *
 * @module render/morph-plan-css
 */
import type { MorphTransitionPlan } from './morph-plan-types';

/**
 * Characters that must never appear raw in a CSS string: the quote and
 * backslash that end or escape it, the angle brackets that would let a value
 * close the enclosing `<style>` element when the stylesheet is injected as
 * markup, `&`, the apostrophe, and control characters (a newline is invalid
 * inside a CSS string).
 */
const CSS_UNSAFE_CHAR = /["\\<>&'\u0000-\u001f\u007f]/u;

/**
 * Escape a value for use inside a double-quoted CSS attribute selector.
 * Unsafe characters become CSS hex escapes (`\3c `), which the stylesheet
 * parser reads back as the same character while keeping the text free of
 * markup. Element ids come from slide part paths, so they are not trusted.
 */
function cssAttributeValue(value: string): string {
	let out = '';
	for (const ch of value) {
		out += CSS_UNSAFE_CHAR.test(ch) ? `\\${(ch.codePointAt(0) ?? 0).toString(16)} ` : ch;
	}
	return out;
}

/**
 * Render a plan as a self-contained stylesheet that drives the morph purely
 * through CSS, with no per-element prop plumbing.
 *
 * Every binding already stamps `data-element-id` on each rendered element, so
 * scoping rules to an ancestor carrying `scopeAttribute` is enough to animate
 * a whole slide. This is what lets Vue / Angular / Svelte / Vanilla play a
 * morph without threading an animation map through their slide-stage
 * components.
 *
 * @param plan - The plan to render.
 * @param scopeAttribute - Attribute (already present on the ancestor) that
 *   scopes these rules, e.g. `data-pptx-morph-incoming`. Pass an empty string
 *   to emit document-level rules: element ids embed their slide path, so they
 *   are unique to the slide being animated and need no ancestor to disambiguate.
 *   That is what lets a binding whose incoming slide is rendered OUTSIDE the
 *   overlay (Angular, React) still drive it from here.
 * @param which - Which half to emit: the live stage's `incoming` elements, the
 *   overlay's `outgoing` ghosts, or the `lifted` copies the overlay paints over
 *   those ghosts (see {@link MorphTransitionPlan.overlayIncomingElements}).
 * @returns Keyframes plus the scoped `animation` rules, ready to inject.
 */
export function buildMorphScopedCss(
	plan: MorphTransitionPlan,
	scopeAttribute: string,
	which: 'incoming' | 'outgoing' | 'lifted' = 'incoming',
): string {
	return `${plan.keyframesCss}\n${buildMorphAnimationRules(plan, scopeAttribute, which)}`;
}

/**
 * Just the `animation` rules of a plan, with no `@keyframes` block.
 *
 * For a binding that already injects {@link MorphTransitionPlan.keyframesCss}
 * itself and applies the element-level animations some other way (React merges
 * them into its per-element animation state), but still needs the descendant
 * rules that {@link MorphTransitionPlan.incomingImageAnimations} cannot be
 * expressed as an inline style for.
 *
 * @param plan - The plan to render.
 * @param scopeAttribute - As {@link buildMorphScopedCss}.
 * @param which - Which half of the transition to emit.
 * @param only - `'image'` emits only the `<img>` descendant rules; omit for all.
 * @returns The newline-joined rules (no trailing newline); `''` when there are none.
 */
export function buildMorphAnimationRules(
	plan: MorphTransitionPlan,
	scopeAttribute: string,
	which: 'incoming' | 'outgoing' | 'lifted' = 'incoming',
	only?: 'image',
): string {
	const prefix = scopeAttribute ? `[${scopeAttribute}] ` : '';
	const rules: string[] = [];
	const emit = (animations: Map<string, string>, suffix: string): void => {
		for (const [elementId, animation] of animations) {
			rules.push(
				`${prefix}[data-element-id="${cssAttributeValue(elementId)}"]${suffix} { animation: ${animation}; }`,
			);
		}
	};
	// `lifted` is the incoming half painted in the overlay rather than on the
	// stage, so it shares the incoming img channel and differs only in which
	// container animation it carries.
	if (only !== 'image') {
		emit(
			which === 'outgoing'
				? plan.outgoingAnimations
				: which === 'lifted'
					? plan.overlayIncomingAnimations
					: plan.incomingAnimations,
			'',
		);
	}
	// The picture-crop channel targets the `<img>` the element renders, which
	// every binding draws inside the `data-element-id` container.
	emit(which === 'outgoing' ? plan.outgoingImageAnimations : plan.incomingImageAnimations, ' img');
	const outline = buildMorphOutlineRules(plan, 'data-element-id', prefix);
	if (outline) {
		rules.push(outline);
	}
	return rules.join('\n');
}

/**
 * The outline-path rules of a plan (see
 * {@link MorphTransitionPlan.outlineAnimations}), keyed on `keyAttribute`.
 *
 * {@link buildMorphAnimationRules} keys them on `data-element-id`; a binding
 * whose overlay paints its halves without that attribute (React's overlay
 * wraps each one in a `data-pptx-morph-outgoing` / `data-pptx-morph-lifted`
 * element named after its id instead) passes that attribute. A re-pathed
 * outline animates the CSS `d` property, so the rules sit behind a support
 * check; without it the static path stays, scaled with its box as before.
 *
 * @returns The rules, or `''` when the plan re-paths no outline.
 */
export function buildMorphOutlineRules(
	plan: MorphTransitionPlan,
	keyAttribute: string,
	prefix = '',
): string {
	const rules = [
		...[...(plan.outlineAnimations ?? [])].map(
			([elementId, animation]) =>
				`${prefix}[${keyAttribute}="${cssAttributeValue(elementId)}"] [data-pptx-outline] path { animation: ${animation}; }`,
		),
		...[...(plan.outlineStrandAnimations ?? [])].flatMap(([elementId, animations]) =>
			animations.map(
				(animation, strand) =>
					`${prefix}[${keyAttribute}="${cssAttributeValue(elementId)}"] [data-pptx-outline] path:nth-of-type(${strand + 1}) { animation: ${animation}; }`,
			),
		),
		...[...(plan.customGeometryAnimations ?? [])].map(
			([elementId, animation]) =>
				`${prefix}[${keyAttribute}="${cssAttributeValue(elementId)}"] [data-pptx-custom-geometry] path { animation: ${animation}; }`,
		),
	];
	const lines = [...(plan.lineAnimations ?? [])].map(
		([elementId, animation]) =>
			`${prefix}[${keyAttribute}="${cssAttributeValue(elementId)}"] [data-pptx-line] { animation: ${animation}; }`,
	);
	const supported = rules.length > 0 ? `@supports (d: path('M0 0')) {\n${rules.join('\n')}\n}` : '';
	return [supported, ...lines].filter(Boolean).join('\n');
}
