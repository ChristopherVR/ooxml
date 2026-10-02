import { getIcon } from 'ooxml-ui';
import { ICON_ACCENTS, ICON_PATHS } from './icon-paths';

const NS = 'http://www.w3.org/2000/svg';
const STROKE = 1.7;

/** Adds or replaces a 24x24 line icon for ribbon commands. Returns false for an empty name or path. */
export function registerRibbonIcon(name: string, path: string): boolean {
	if (!name || !path) return false;
	ICON_PATHS[name] = path;
	return true;
}

export function hasRibbonIcon(name: string | undefined): boolean {
	return Boolean(name && (name in ICON_PATHS || getIcon(name)));
}

/**
 * An `<svg>` for `name` at `size` CSS pixels: a ribbon icon (24-unit grid) or, failing that, an
 * ooxml-ui icon (20-unit grid, registered with `registerIcon`). Unknown names draw nothing.
 */
export function ribbonIcon(doc: Document, name: string | undefined, size = 16): SVGSVGElement {
	const svg = doc.createElementNS(NS, 'svg');
	svg.setAttribute('width', String(size));
	svg.setAttribute('height', String(size));
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('ribbon-icon');
	if (!name) return svg;
	svg.dataset.icon = name;
	const accent = ICON_ACCENTS[name];
	if (accent) svg.dataset.accent = accent;
	const local = ICON_PATHS[name];
	const shared = local ? undefined : getIcon(name);
	svg.setAttribute('viewBox', local ? '0 0 24 24' : (shared?.viewBox ?? '0 0 20 20'));
	const d = local ?? shared?.d;
	if (!d) return svg;
	const path = doc.createElementNS(NS, 'path');
	path.setAttribute('d', d);
	path.setAttribute('fill', 'none');
	path.setAttribute('stroke', 'currentColor');
	path.setAttribute('stroke-width', String(local ? STROKE : 1.5));
	path.setAttribute('stroke-linecap', 'round');
	path.setAttribute('stroke-linejoin', 'round');
	svg.append(path);
	return svg;
}
