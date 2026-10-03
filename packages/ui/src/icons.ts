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
	cut: 'M7 3l7 10M13 3 6 13M3 15.5a2 2 0 1 0 4 0 2 2 0 1 0-4 0M13 15.5a2 2 0 1 0 4 0 2 2 0 1 0-4 0',
	paste: 'M7 3.5h6v2.5H7ZM7 4.5H4.5v13h11v-13H13',
	formatPainter: 'M4 3h10v4H4ZM14 5h2.5v4.5H9.5V12M8.5 12h2v5h-2Z',
	bold: 'M6 3.5h4.5a3 3 0 0 1 0 6H6ZM6 9.5h5.5a3.5 3.5 0 0 1 0 7H6Z',
	italic: 'M9 3.5h6M5 16.5h6M12 3.5 8 16.5',
	underline: 'M6 3v6a4 4 0 0 0 8 0V3M5 17h10',
	strikethrough:
		'M3.5 10h13M13.5 5.5C13 4.3 11.7 3.5 10 3.5c-2 0-3.5 1-3.5 2.6 0 1.4 1.2 2.2 3.5 2.8M6.5 13.8c.5 1.6 1.9 2.7 3.8 2.7 2 0 3.5-1.1 3.5-2.8',
	fontColor: 'M5.5 13 10 3l4.5 10M7.2 9.3h5.6M3 17h14',
	growFont: 'M2.5 16 6.5 6l4 10M4 12.3h5M13 8.5l2-3 2 3M15 5.5V12',
	shrinkFont: 'M2.5 16 6.5 6l4 10M4 12.3h5M13 9l2 3 2-3M15 12V5.5',
	changeCase:
		'M2 15 5.5 5 9 15M3.2 11.7h4.6M15.5 9.5v5.5M15.5 12.2a2.3 2.3 0 1 1-4.6 0 2.3 2.3 0 1 1 4.6 0',
	alignLeft: 'M3 4.5h14M3 8.5h9M3 12.5h14M3 16.5h9',
	alignCenter: 'M3 4.5h14M5.5 8.5h9M3 12.5h14M5.5 16.5h9',
	alignRight: 'M3 4.5h14M8 8.5h9M3 12.5h14M8 16.5h9',
	justify: 'M3 4.5h14M3 8.5h14M3 12.5h14M3 16.5h14',
	alignTop: 'M3 3.5h14M6 7h8M6 10.5h8',
	alignMiddle: 'M3 10h2M15 10h2M6 6.5h8M6 13.5h8M6 10h8',
	alignBottom: 'M3 16.5h14M6 9.5h8M6 13h8',
	bullets: 'M8 5h9M8 10h9M8 15h9M3.5 5h1M3.5 10h1M3.5 15h1',
	indentIncrease: 'M3 4h14M9 8h8M9 12h8M3 16h14M3 7.5 6 10l-3 2.5',
	indentDecrease: 'M3 4h14M9 8h8M9 12h8M3 16h14M6 7.5 3 10l3 2.5',
	fill: 'M3.5 10 9 4.5l6 6L9.5 16ZM9 4.5 7 2.5M16.5 12.5s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3',
	line: 'M3.5 16.5 16.5 3.5',
	effects: 'M10 2.5l1.9 4.6 4.6 1.9-4.6 1.9L10 15.5l-1.9-4.6L3.5 9l4.6-1.9Z',
	quickStyles:
		'M10 3a7 7 0 1 0 0 14c1 0 1.5-.7 1.5-1.5S11 14 11 13s.8-1.5 2-1.5h2a2 2 0 0 0 2-2A7 7 0 0 0 10 3ZM6.5 9.5h.1M8.5 6.2h.1M12.5 6.5h.1',
	bringToFront: 'M7.5 7.5h9v9h-9ZM12.5 7.5v-4h-9v9h4',
	sendToBack: 'M3.5 3.5h9v9h-9ZM12.5 7.5h4v9h-9v-4',
	group: 'M3 3h5v5H3ZM12 12h5v5h-5ZM2 2h1M17 2h1M2 18h1M17 18h1',
	alignObjects: 'M3 2.5v15M6 5h8v4H6ZM6 11h11v4H6Z',
	position:
		'M10 2.5v15M2.5 10h15M7.5 5 10 2.5 12.5 5M7.5 15l2.5 2.5 2.5-2.5M5 7.5 2.5 10 5 12.5M15 7.5l2.5 2.5-2.5 2.5',
	connector: 'M3.5 15.5h5v-11h8M14.5 2.5l2 2-2 2M2.5 13.5v4',
	textBox: 'M3 3h14v14H3ZM6.5 6.5h7M10 6.5v7',
	ruler: 'M2 7h16v6H2ZM5 7v3M8 7v2M11 7v3M14 7v2',
	launcher: 'M5 15 15 5M9 5h6v6',
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
