import {
	cloneElementForPaste,
	copyElementsToSystemClipboard,
	isTemplateElementId,
} from 'ooxml-ui/pptx';
import type { ElementNodeRasterizer } from 'ooxml-ui/pptx';
/**
 * useElementClipboard: the in-memory single-element copy / cut / paste buffer.
 *
 * The paste source stays this in-memory buffer: PowerPoint pastes a shape with
 * all its OOXML fidelity intact, which a text/image clipboard round-trip would
 * flatten. Copy also offers other applications a picture and the text of the
 * selection through the system clipboard (shared `copyElementsToSystemClipboard`).
 * Cloning, descendant re-id and the paste cascade offset go through the shared
 * clipboard codec so this matches paste/duplicate/ungroup everywhere else.
 */
import { cloneElement } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';
import type { ComputedRef, Ref } from 'vue';
import { computed, shallowRef } from 'vue';

export interface UseElementClipboardOptions {
	/** Resolve an element on the active slide (slide content only, not template). */
	findSlideElement: (id: string) => PptxElement | undefined;
	addElement: (element: PptxElement) => void;
	removeElement: (id: string) => void;
	selectedElementIds: Ref<string[]>;
	/** Renders a copied shape for the system clipboard's picture (optional). */
	rasterize?: ElementNodeRasterizer;
}

export interface UseElementClipboardResult {
	clipboard: Ref<PptxElement | null>;
	hasClipboard: ComputedRef<boolean>;
	copyElement: (id: string) => void;
	cutElement: (id: string) => void;
	/** Returns the freshly-inserted clone, or undefined when the clipboard was empty. */
	pasteElement: () => PptxElement | undefined;
}

export function useElementClipboard(
	options: UseElementClipboardOptions,
): UseElementClipboardResult {
	// shallowRef: a deep-reactive ref() wraps an assigned object in a Proxy,
	// which structuredClone (used by cloneElementForPaste) cannot clone. The
	// clipboard buffer is never rendered field-by-field, so it does not need
	// Vue's deep reactivity.
	const clipboard = shallowRef<PptxElement | null>(null);
	const hasClipboard = computed(() => clipboard.value !== null);

	function copyElement(id: string): void {
		const el = options.findSlideElement(id);
		if (el) {
			clipboard.value = cloneElement(el);
			// Other applications get a picture of the selection and its text (best effort).
			const ids = options.selectedElementIds.value.includes(id)
				? options.selectedElementIds.value
				: [id];
			const elements = ids
				.map((value) => options.findSlideElement(value))
				.filter((value): value is PptxElement => Boolean(value));
			void copyElementsToSystemClipboard({
				elements,
				...(options.rasterize ? { rasterize: options.rasterize } : {}),
			});
		}
	}

	function cutElement(id: string): void {
		copyElement(id);
		options.removeElement(id);
		options.selectedElementIds.value = options.selectedElementIds.value.filter((x) => x !== id);
	}

	function pasteElement(): PptxElement | undefined {
		if (!clipboard.value) {
			return undefined;
		}
		const copy = cloneElementForPaste(clipboard.value, {
			intoTemplate: isTemplateElementId(clipboard.value.id),
		});
		options.addElement(copy);
		options.selectedElementIds.value = [copy.id];
		return copy;
	}

	return { clipboard, hasClipboard, copyElement, cutElement, pasteElement };
}
