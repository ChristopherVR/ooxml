/**
 * Classify the element a key press came from, for {@link ./editor-keymap}.
 *
 * Split out of `editor-keymap.ts` (file-size limit) and kept together because
 * both predicates answer the same question for every binding: "is this key
 * press aimed at the editor, or at some other control that owns it?"
 *
 * @module render/editor-key-target
 */

const FORM_FIELD_TAGS = /^(?:INPUT|TEXTAREA|SELECT)$/u;

/** Minimal element shape the predicates read; `unknown` targets are narrowed to it. */
interface KeyTargetElement {
	tagName?: string;
	isContentEditable?: boolean;
	closest?: (selector: string) => unknown;
	shadowRoot?: { activeElement?: unknown } | null;
}

function asElement(target: unknown): (KeyTargetElement & { tagName: string }) | null {
	const element = target as KeyTargetElement | null;
	if (!element || typeof element.tagName !== 'string') {
		return null;
	}
	return element as KeyTargetElement & { tagName: string };
}

/**
 * The element a key press really came from. A custom element with a shadow root
 * (the SmartArt text pane) retargets the event to its host, so
 * a field inside it looks like the host; follow the focused element down through
 * open shadow roots to the real field.
 */
function deepTarget(target: unknown): unknown {
	let current = target as KeyTargetElement | null;
	for (let depth = 0; depth < 8; depth++) {
		const inner = current?.shadowRoot?.activeElement as KeyTargetElement | null | undefined;
		if (!inner || typeof inner.tagName !== 'string') {
			break;
		}
		current = inner;
	}
	return current;
}

/**
 * True when a key press is the user typing into a field rather than driving the
 * editor. Kept here so every binding classifies the same targets: a binding that
 * forgot `<select>` would swallow the arrow keys of its own dropdowns.
 */
export function isEditorTextInputTarget(target: unknown): boolean {
	const element = asElement(deepTarget(target));
	if (!element) {
		return false;
	}
	return FORM_FIELD_TAGS.test(element.tagName) || element.isContentEditable === true;
}
