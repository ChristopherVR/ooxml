import { html, type PropertyValues } from 'lit';
import type { DiagramDrawing } from 'ooxml-core/diagram';
import { OfficeElement, controlStyles } from '../base.js';
import { definer } from '../registry.js';
import {
	renderDiagramDrawing,
	type SchemeColors,
	type SchemeFonts,
	type SmartArtRenderReport,
} from './smartart-svg.js';
import css from './smartart.css?raw';

export type { SchemeColors, SchemeFonts, SmartArtRenderReport } from './smartart-svg.js';
export { renderDiagramDrawing } from './smartart-svg.js';
export type OfficeSmartArtRenderEvent = CustomEvent<SmartArtRenderReport>;

/**
 * `<office-ui-smartart>` draws the cached drawing of a SmartArt diagram (core
 * `DiagramDrawing`, from `ooxml-core/diagram`) as SVG. Set `drawing` (and
 * optionally `schemeColors`, resolved theme colours such as `{ accent1: '#4472c4' }`,
 * and `schemeFonts`, the theme's `major`/`minor` Latin typefaces);
 * `label` becomes the accessible name (`role="img"`). It never re-lays-out a diagram: it shows
 * what the producing application last computed. Approximations (preset outlines drawn as
 * rectangles, gradient and pattern fills, 3D) are reported honestly in the
 * `office-smartart-render` event detail and in `data-approximated` attributes, never hidden.
 */
export class OfficeUiSmartArt extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		drawing: { attribute: false },
		schemeColors: { attribute: false },
		schemeFonts: { attribute: false },
		label: { type: String },
	};
	declare drawing: DiagramDrawing | undefined;
	declare schemeColors: SchemeColors;
	declare schemeFonts: SchemeFonts;
	declare label: string | null;
	private figure: SVGSVGElement | undefined;
	private lastReport: SmartArtRenderReport | undefined;

	constructor() {
		super();
		this.drawing = undefined;
		this.schemeColors = {};
		this.schemeFonts = {};
		this.label = null;
	}

	/** The report of the last render (undefined while empty). */
	get report(): SmartArtRenderReport | undefined {
		return this.lastReport;
	}

	/** The SVG is rebuilt only when the drawing or the scheme colours change. */
	protected override willUpdate(changed: PropertyValues<this>): void {
		if (changed.has('drawing') || changed.has('schemeColors') || changed.has('schemeFonts')) {
			if (!this.drawing || this.drawing.shapes.length === 0) {
				this.figure = undefined;
				this.lastReport = undefined;
			} else {
				const { svg, report } = renderDiagramDrawing(
					this.ownerDocument,
					this.drawing,
					this.schemeColors,
					this.schemeFonts,
				);
				svg.setAttribute('role', 'img');
				this.figure = svg;
				this.lastReport = report;
			}
		}
		this.figure?.setAttribute('aria-label', this.label ?? 'Diagram');
	}

	protected override updated(changed: PropertyValues<this>): void {
		if (
			this.lastReport &&
			(changed.has('drawing') || changed.has('schemeColors') || changed.has('schemeFonts'))
		)
			this.fire('office-smartart-render', this.lastReport);
	}

	protected override render() {
		if (this.figure) return this.figure;
		return html`<div class="empty"
			>${this.drawing ? 'This diagram has no cached drawing to show.' : ''}</div
		>`;
	}
}

export const defineSmartArt = definer('office-ui-smartart', () => OfficeUiSmartArt);
