/**
 * Whether this viewer draws the deck's linked web content (the shared
 * `resolveExternalContentView(...).allowed`), provided once by
 * `PowerPointViewer` and read by every surface that draws the model
 * (`SlideStage`, the canvas background, table styles, inspector previews).
 * Enable Content changes the value and those surfaces redraw: nothing reloads
 * and the model keeps its links. Outside a provider, linked content is
 * withheld (fail closed). Mirrors React's `LinkedContentContext`.
 */
import { linkedContentUrl, linkedContentView } from 'ooxml-ui/pptx';
import type { ComputedRef, InjectionKey, MaybeRefOrGetter } from 'vue';
import { computed, inject, provide, toValue } from 'vue';

/** Typed injection key for the viewer's linked-content decision. */
export const LinkedContentKey: InjectionKey<MaybeRefOrGetter<boolean>> =
	Symbol('pptx-vue-linked-content');

/** Provide `allowed` to the components below. */
export function provideLinkedContent(allowed: MaybeRefOrGetter<boolean>): void {
	provide(LinkedContentKey, allowed);
}

/**
 * Provide `allowed()` below this component when it is set, else keep the surrounding
 * decision: a building block a host shell renders takes it as its `linkedContent` prop.
 * Returns the decision now in force, for the component's own drawing.
 */
export function provideLinkedContentScope(
	allowed: () => boolean | undefined,
): ComputedRef<boolean> {
	const parent = inject(LinkedContentKey, false);
	const scoped = computed(() => allowed() ?? toValue(parent));
	provide(LinkedContentKey, scoped);
	return scoped;
}

/** The surrounding decision, `false` outside a provider. */
export function useLinkedContentAllowed(): ComputedRef<boolean> {
	const source = inject(LinkedContentKey, false);
	return computed(() => toValue(source));
}

/** `value()` as the surfaces under this component may draw it (`linkedContentView`). */
export function useLinkedContentView<T>(value: () => T): ComputedRef<T> {
	const allowed = useLinkedContentAllowed();
	return computed(() => linkedContentView(value(), allowed.value));
}

/**
 * One address a preview draws (an inspector thumbnail), withheld while it is a web address
 * this viewer has not allowed. Editing code keeps reading the model itself.
 */
export function useLinkedContentUrl(
	url: () => string | undefined,
): ComputedRef<string | undefined> {
	const allowed = useLinkedContentAllowed();
	return computed(() => linkedContentUrl(url(), allowed.value));
}
