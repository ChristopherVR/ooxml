import { SHAPE_THUMB_VIEW_BOX, homeLabel } from '../render';
import type { RibbonHomeItem, RibbonHomeViewState } from '../render';
import { INSERT_SHAPE_GLYPH_PATHS, insertGlyphTransform } from './ribbon-insert-icons';

const SVG = 'http://www.w3.org/2000/svg';

function shapeGlyph(doc: Document, icon: string): SVGSVGElement | undefined {
	const [, spec] = icon.split(':');
	const [glyph, glyphClass = ''] = (spec ?? '').split('|');
	const path = INSERT_SHAPE_GLYPH_PATHS[glyph];
	if (!path) {
		return undefined;
	}
	const svg = doc.createElementNS(SVG, 'svg');
	svg.setAttribute('viewBox', '0 0 16 16');
	svg.setAttribute('aria-hidden', 'true');
	svg.style.transform = insertGlyphTransform(glyphClass);
	const shape = doc.createElementNS(SVG, 'path');
	shape.setAttribute('d', path);
	svg.append(shape);
	return svg;
}

function thumbGlyph(doc: Document, path: string): SVGSVGElement {
	const svg = doc.createElementNS(SVG, 'svg');
	svg.setAttribute('viewBox', SHAPE_THUMB_VIEW_BOX);
	svg.setAttribute('aria-hidden', 'true');
	const shape = doc.createElementNS(SVG, 'path');
	shape.setAttribute('d', path);
	svg.append(shape);
	return svg;
}

/**
 * Paint menu rows into `menu`. A row marked `checked` (or equal to `value`)
 * is exposed as a checked radio item; headings appear once per group.
 */
export function paintHomeMenu(
	doc: Document,
	menu: HTMLElement,
	items: readonly RibbonHomeItem[],
	state: RibbonHomeViewState,
	value: string | number | undefined,
	pick: (value: string) => void,
	grid = false,
): void {
	menu.classList.toggle('grid', grid);
	const rows: HTMLElement[] = [];
	let group: string | undefined;
	// A thumbnail gallery picks a shape to draw; it has no current choice to check.
	const checkable =
		!grid && (value !== undefined || items.some((row) => row.checked !== undefined));
	for (const row of items) {
		const heading = row.groupKey ? homeLabel(state, row.groupKey, row.group ?? '') : row.group;
		if (heading !== undefined && heading !== group) {
			const title = doc.createElement('div');
			title.className = 'heading';
			title.setAttribute('role', 'presentation');
			title.textContent = heading;
			rows.push(title);
		}
		group = heading;
		if (row.separator) {
			const rule = doc.createElement('hr');
			rule.setAttribute('role', 'separator');
			rows.push(rule);
		}
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'item';
		button.dataset.value = row.value;
		button.disabled = Boolean(row.disabled);
		const label = row.label ?? homeLabel(state, row.labelKey ?? '', row.fallback ?? row.value);
		const checked = row.checked ?? (value !== undefined && String(value) === row.value);
		button.setAttribute('role', checkable ? 'menuitemradio' : 'menuitem');
		if (checkable) {
			button.setAttribute('aria-checked', String(checked));
		}
		for (const [name, attr] of Object.entries(row.attrs ?? {})) {
			button.setAttribute(name, attr);
		}
		const glyph = row.thumb
			? thumbGlyph(doc, row.thumb)
			: row.icon?.startsWith('shape:')
				? shapeGlyph(doc, row.icon)
				: undefined;
		if (glyph) {
			button.append(glyph);
		}
		if (grid && row.thumb) {
			// A gallery cell: the name is its ScreenTip and accessible name, not visible text.
			const reason = row.disabled && (row.descriptionKey || row.description);
			button.classList.add('cell');
			button.setAttribute('aria-label', label);
			button.title = reason
				? `${label}
${homeLabel(state, row.descriptionKey ?? '', row.description ?? '')}`
				: label;
			// Keep the text selection and caret in the slide while a menu row is pressed.
			button.addEventListener('mousedown', (event) => event.preventDefault());
			button.addEventListener('click', () => pick(row.value));
			rows.push(button);
			continue;
		}
		const text = doc.createElement('span');
		text.className = 'label';
		text.textContent = label;
		if (row.fontFamily) {
			text.style.fontFamily = row.fontFamily;
		}
		button.append(text);
		if (checked) {
			const mark = doc.createElement('span');
			mark.className = 'mark';
			mark.setAttribute('aria-hidden', 'true');
			mark.textContent = '\u2022';
			button.append(mark);
		}
		// Keep the text selection and caret in the slide while a menu row is pressed.
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', () => pick(row.value));
		rows.push(button);
	}
	menu.replaceChildren(...rows);
}
