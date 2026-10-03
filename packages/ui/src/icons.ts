import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/** A trusted icon: SVG path data in a 20x20 box. Hosts never inject markup. */
export interface OfficeIcon {
	d: string;
	/** Defaults to `0 0 20 20`. */
	viewBox?: string;
}

const registry = new Map<string, OfficeIcon>();

/**
 * Neutral icons every Office product needs. Derived from pptx-viewer's ribbon icon set
 * (packages/shared/src/web-components/ribbon-icons.ts); product-specific glyphs stay there.
 */
const BUILT_IN: Readonly<Record<string, string>> = {
	check: 'm3 10 4.5 4.5L17 5',
	close: 'M5 5l10 10M15 5 5 15',
	chevronDown: 'm5 8 5 5 5-5',
	chevronLeft: 'm12 4-6 6 6 6',
	chevronRight: 'm8 4 6 6-6 6',
	copy: 'M7 7h10v10H7ZM3 13V3h10',
	trash: 'M4 6h12M8 6V4h4v2M6 6l.8 10h6.4L14 6',
	lock: 'M4 9h12v8H4ZM7 9V6.5a3 3 0 0 1 6 0V9',
	eyeOff: 'M2.5 10S5.5 5 10 5s7.5 5 7.5 5-3 5-7.5 5-7.5-5-7.5-5zM4 4l12 12',
	message: 'M3 4h14v9H8l-5 4z',
	pencil: 'M12 3l5 5M3 17l1-5L13 3l4 4-9 9Z',
	help: 'M3 10a7 7 0 1 0 14 0 7 7 0 1 0-14 0M8 7a2 2 0 1 1 3 1.8c-.7.3-1 .8-1 1.7M10 14v.1',
	settings: 'M3 5h14M6 10h8M8 15h4M6 3v4M12 8v4M10 13v4',
	clock: 'M3 10a7 7 0 1 0 14 0 7 7 0 1 0-14 0M10 6v4l3 2',
	reset: 'M4 10a6 6 0 1 1 1.8 4.2M4 6v4h4',
	search: 'M3 8.5a5.5 5.5 0 1 0 11 0 5.5 5.5 0 1 0-11 0M13 13l4 4',
	undo: 'M7.5 11.5 3.5 7.5l4-4M3.5 7.5h9a4 4 0 0 1 0 8H10',
	redo: 'm12.5 11.5 4-4-4-4M16.5 7.5h-9a4 4 0 0 0 0 8H10',
	pointer: 'm5 2.5v13.5l3.6-3.6 2.4 5.1 2.1-1-2.4-5H15.5Z',
	rectangle: 'M3 5.5h14v9H3Z',
	fullscreen: 'M3 7.5V3h4.5M12.5 3H17v4.5M17 12.5V17h-4.5M7.5 17H3v-4.5',
	grid: 'M3 3h14v14H3ZM7.7 3v14M12.3 3v14M3 7.7h14M3 12.3h14',
	fitPage: 'M6.5 6.5h7v7h-7ZM2.5 6V2.5H6M14 2.5h3.5V6M17.5 14v3.5H14M6 17.5H2.5V14',
	pageWidth: 'M5.5 3.5h9v13h-9ZM1.5 10h17M4 8l-2.5 2L4 12M16 8l2.5 2-2.5 2',
	zoomIn: 'M3 8.5a5.5 5.5 0 1 0 11 0 5.5 5.5 0 1 0-11 0M13 13l4 4M6 8.5h5M8.5 6v5',
	zoomOut: 'M3 8.5a5.5 5.5 0 1 0 11 0 5.5 5.5 0 1 0-11 0M13 13l4 4M6 8.5h5',
};
for (const [name, d] of Object.entries(BUILT_IN)) registry.set(name, { d });

/** Register or replace an icon. Returns false for an empty name or path. */
export function registerIcon(name: string, icon: OfficeIcon | string): boolean {
	const value = typeof icon === 'string' ? { d: icon } : icon;
	if (!name || !value.d) return false;
	registry.set(name, value);
	return true;
}

export function getIcon(name: string | null | undefined): OfficeIcon | undefined {
	return name ? registry.get(name) : undefined;
}

export function listIcons(): string[] {
	return [...registry.keys()].sort();
}

const SVG_NS = 'http://www.w3.org/2000/svg';

export function createIconSvg(doc: Document): SVGSVGElement {
	const svg = doc.createElementNS(SVG_NS, 'svg');
	svg.setAttribute('aria-hidden', 'true');
	return svg;
}

/** Fill an `<svg>` with a registered icon (empty when unknown). Used by every control. */
export function paintIcon(svg: SVGElement, name: string | null | undefined): boolean {
	const icon = getIcon(name);
	svg.setAttribute('viewBox', icon?.viewBox ?? '0 0 20 20');
	svg.replaceChildren();
	if (!icon) return false;
	const path = svg.ownerDocument.createElementNS(SVG_NS, 'path');
	path.setAttribute('d', icon.d);
	svg.append(path);
	return true;
}

export const ICON_CSS = `
svg { width: 1em; height: 1em; fill: none; stroke: currentColor; stroke-width: 1.5;
	stroke-linecap: round; stroke-linejoin: round; }
`;

/** `<office-ui-icon name="check" label="Done">`: decorative unless `label` is set. */
export const defineIcon = definer('office-ui-icon', () => {
	class OfficeUiIcon extends HTMLElement {
		static observedAttributes = ['name', 'label'];
		private readonly svg: SVGSVGElement;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(
				root,
				controlCss(`:host { display: inline-flex; width: 1em; height: 1em; } ${ICON_CSS}`),
			);
			this.svg = createIconSvg(this.ownerDocument);
			root.append(this.svg);
		}
		connectedCallback(): void {
			this.sync();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get name(): string {
			return this.getAttribute('name') ?? '';
		}
		set name(value: string) {
			this.setAttribute('name', value);
		}
		private sync(): void {
			paintIcon(this.svg, this.name);
			const label = this.getAttribute('label');
			if (label) {
				this.setAttribute('role', 'img');
				this.setAttribute('aria-label', label);
				this.removeAttribute('aria-hidden');
			} else {
				this.removeAttribute('role');
				this.removeAttribute('aria-label');
				this.setAttribute('aria-hidden', 'true');
			}
		}
	}
	return OfficeUiIcon;
});
