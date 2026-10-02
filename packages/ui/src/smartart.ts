import type { DiagramDrawing } from 'ooxml-core/diagram';
import { definer, emit } from './registry.js';
import {
	renderDiagramDrawing,
	type SchemeColors,
	type SmartArtRenderReport,
} from './smartart-svg.js';
import { attachStyles, controlCss } from './styles.js';

export type { SchemeColors, SmartArtRenderReport } from './smartart-svg.js';
export { renderDiagramDrawing } from './smartart-svg.js';
export type OfficeSmartArtRenderEvent = CustomEvent<SmartArtRenderReport>;

const CSS = `
:host { display: block; }
svg { display: block; width: 100%; height: auto; max-height: 100%; }
text { font-family: var(--office-font, system-ui, sans-serif); pointer-events: none; }
.empty { padding: 8px; font-size: 12px; color: var(--office-muted-foreground, #6b7280); }
@media (forced-colors: active) { text { fill: CanvasText; } }
`;

/**
 * `<office-ui-smartart>` draws the cached drawing of a SmartArt diagram (core
 * `DiagramDrawing`, from `ooxml-core/diagram`) as SVG. Set `drawing` (and
 * optionally `schemeColors`, resolved theme colours such as `{ accent1: '#4472c4' }`);
 * `label` becomes the accessible name (`role="img"`). It never re-lays-out a diagram: it shows
 * what the producing application last computed. Approximations (preset outlines drawn as
 * rectangles, gradient and pattern fills, 3D) are reported honestly in the
 * `office-smartart-render` event detail and in `data-approximated` attributes, never hidden.
 */
export const defineSmartArt = definer('office-ui-smartart', () => {
	class OfficeUiSmartArt extends HTMLElement {
		static observedAttributes = ['label'];
		private model: DiagramDrawing | undefined;
		private scheme: SchemeColors = {};
		private lastReport: SmartArtRenderReport | undefined;
		private readonly root: ShadowRoot;
		constructor() {
			super();
			this.root = this.attachShadow({ mode: 'open' });
			attachStyles(this.root, controlCss(CSS));
		}
		connectedCallback(): void {
			this.render();
		}
		attributeChangedCallback(): void {
			this.render();
		}
		get drawing(): DiagramDrawing | undefined {
			return this.model;
		}
		set drawing(next: DiagramDrawing | undefined) {
			this.model = next;
			this.render();
		}
		get schemeColors(): SchemeColors {
			return this.scheme;
		}
		set schemeColors(next: SchemeColors) {
			this.scheme = next;
			this.render();
		}
		/** The report of the last render (undefined while empty). */
		get report(): SmartArtRenderReport | undefined {
			return this.lastReport;
		}
		private render(): void {
			const doc = this.ownerDocument;
			for (const el of this.root.querySelectorAll('svg, .empty')) el.remove();
			const label = this.getAttribute('label') ?? 'Diagram';
			if (!this.model || this.model.shapes.length === 0) {
				this.lastReport = undefined;
				const empty = doc.createElement('div');
				empty.className = 'empty';
				empty.textContent = this.model ? 'This diagram has no cached drawing to show.' : '';
				this.root.append(empty);
				return;
			}
			const { svg, report } = renderDiagramDrawing(doc, this.model, this.scheme);
			svg.setAttribute('role', 'img');
			svg.setAttribute('aria-label', label);
			this.root.append(svg);
			this.lastReport = report;
			emit(this, 'office-smartart-render', report);
		}
	}
	return OfficeUiSmartArt;
});
