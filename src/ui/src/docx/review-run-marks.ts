import { DOMSerializer } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
import type { MarkViewConstructor } from 'prosemirror-view';
import type { ReviewDisplayMode } from 'ooxml-core/docx';

const FORMAT_MARKS = [
	'bold',
	'italic',
	'underline',
	'strike',
	'highlight',
	'verticalAlign',
	'language',
	'runRtl',
	'font',
	'runProperties',
] as const;
const NEUTRAL_FORMAT =
	'font:inherit;color:inherit;background-color:transparent;text-decoration:none;text-transform:none;font-variant:normal;letter-spacing:normal;vertical-align:baseline;position:static;top:auto;opacity:1;font-kerning:auto';

/**
 * In Original mode, direct marks keep their document identity but defer their appearance to
 * run-style decorations. Neutralizing outer marks prevents underline, highlight and script
 * positioning from leaking into a prior-properties decoration nested inside them.
 */
export function reviewRunMarksPlugin(getMode: () => ReviewDisplayMode): Plugin {
	const refreshers = new Set<() => void>();
	const markViews: Record<string, MarkViewConstructor> = {};
	for (const name of FORMAT_MARKS)
		markViews[name] = (mark, view, inline) => {
			const rendered = DOMSerializer.renderSpec(
				view.dom.ownerDocument,
				mark.type.spec.toDOM!(mark, inline),
			);
			const dom = rendered.dom as HTMLElement;
			const originals = new Map(
				['style', 'class', 'lang', 'dir'].map((name) => [name, dom.getAttribute(name)]),
			);
			const refresh = () => {
				for (const [name, value] of originals)
					if (value === null) dom.removeAttribute(name);
					else dom.setAttribute(name, value);
				if (getMode() !== 'original') return;
				dom.style.cssText = NEUTRAL_FORMAT;
				dom.removeAttribute('lang');
				dom.removeAttribute('dir');
				dom.classList.remove('dve-hidden-text');
			};
			refreshers.add(refresh);
			refresh();
			return {
				dom,
				...(rendered.contentDOM ? { contentDOM: rendered.contentDOM as HTMLElement } : {}),
				ignoreMutation: (mutation) => mutation.type === 'attributes' && mutation.target === dom,
				destroy: () => {
					refreshers.delete(refresh);
				},
			};
		};
	return new Plugin({
		props: { markViews },
		view() {
			let mode = getMode();
			return {
				update() {
					const next = getMode();
					if (next === mode) return;
					mode = next;
					for (const refresh of refreshers) refresh();
				},
				destroy: () => refreshers.clear(),
			};
		},
	});
}
