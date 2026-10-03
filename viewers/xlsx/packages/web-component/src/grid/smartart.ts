// SmartArt graphics over the grid. The core reads the drawing the producing application cached
// (`SmartArtObject.diagram`, EMU relative to the frame) and `<office-ui-smartart>` (ooxml-ui)
// draws it; the layout is never recomputed and the content is not editable, which the object's
// notice says. Without a cached drawing a placeholder lists the diagram's text nodes.
import { THEME_SLOTS, type SmartArtObject, type ThemePalette } from '@christophervr/xlsx-core';
import { defineSmartArt, type SchemeColors } from 'ooxml-ui';
import type { EditorContext } from '../context.js';
import { h } from './dom.js';

const ALIASES: Readonly<Record<string, string>> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
};

/** The workbook theme as the scheme colours the SmartArt drawing refers to (`#rrggbb`). */
export function smartArtSchemeColors(theme: ThemePalette | undefined): SchemeColors {
	const colors: Record<string, string> = {};
	THEME_SLOTS.forEach((slot, index) => {
		const value = theme?.colors[index];
		if (value && /^[0-9a-f]{6}$/i.test(value)) colors[slot] = `#${value.toLowerCase()}`;
	});
	for (const [alias, slot] of Object.entries(ALIASES)) {
		const value = colors[slot];
		if (value) colors[alias] = value;
	}
	return colors;
}

/**
 * Paints a SmartArt frame into `node`. The shared SVG scales to the positioned frame,
 * following both zoom and drawing resize without clipping the original cached extent.
 */
export function paintSmartArt(
	ctx: EditorContext,
	node: HTMLElement,
	obj: SmartArtObject,
	theme: ThemePalette | undefined,
): void {
	const doc = node.ownerDocument;
	node.classList.add('xg-smartart');
	const label = obj.name || ctx.t('SmartArt graphic');
	node.setAttribute('role', 'img');
	node.setAttribute('aria-label', label);
	const notice = ctx.t(obj.notice);
	node.title = notice;
	const diagram = obj.diagram;
	if (diagram && diagram.shapes.length) {
		defineSmartArt(doc.defaultView?.customElements);
		const art = doc.createElement('office-ui-smartart') as HTMLElement & {
			drawing?: SmartArtObject['diagram'];
			schemeColors?: SchemeColors;
		};
		art.setAttribute('label', label);
		art.style.width = '100%';
		art.style.height = '100%';
		art.schemeColors = smartArtSchemeColors(theme);
		art.drawing = diagram;
		node.append(art);
	} else {
		node.classList.add('xg-smartart-placeholder');
		const list = h(doc, 'ul', 'xg-smartart-nodes');
		for (const item of obj.nodes) {
			if (!item.text.trim()) continue;
			const entry = h(doc, 'li');
			entry.textContent = item.text;
			if (item.parentId && obj.nodes.some((n) => n.id === item.parentId && n.text.trim()))
				entry.classList.add('xg-smartart-child');
			list.append(entry);
		}
		node.append(list);
	}
	const badge = h(doc, 'div', 'xg-smartart-notice');
	badge.textContent = ctx.t('SmartArt (display only)');
	badge.title = notice;
	node.append(badge);
}
