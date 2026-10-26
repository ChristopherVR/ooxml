// SmartArt graphics over the grid. The core reads the drawing the producing application cached
// (`SmartArtObject.diagram`, EMU relative to the frame) and `<office-ui-smartart>` (ooxml-ui)
// draws it. A diagram without one (inserted or edited here, or saved without a cache) is drawn
// with the shared approximate layout (`approximateDiagramDrawingFromNodes`, as Word does), never
// Excel's own layout, which the badge and the object's notice say. Editing goes through the
// SmartArt Design tab and the text pane (../smartart-controller.ts).
import { previewDiagramLayout, resolveDiagramDrawingStyles } from 'ooxml-core/diagram';
import { THEME_SLOTS, type SmartArtObject, type ThemePalette } from 'ooxml-core/xlsx';
import { defineSmartArt, type SchemeColors, type SchemeFonts } from '../../index';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { h } from './dom';

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
	size: { w: number; h: number } = {
		w: (obj.anchor.ext?.cx ?? 4572000) / 9525,
		h: (obj.anchor.ext?.cy ?? 2667000) / 9525,
	},
): void {
	const doc = node.ownerDocument;
	node.classList.add('xg-smartart');
	const label = obj.name || ctx.t('SmartArt graphic');
	node.setAttribute('role', 'img');
	node.setAttribute('aria-label', label);
	const notice = ctx.t(obj.notice);
	node.title = notice;
	const cached = obj.diagram?.shapes.length ? obj.diagram : undefined;
	const preview = !cached
		? previewDiagramLayout({
				extentEmu: { width: Math.round(size.w * 9525), height: Math.round(size.h * 9525) },
				nodes: obj.nodes,
				layoutId: obj.layoutId,
				family: obj.layoutFamily,
				colorsId: obj.colorsId,
				source: {
					...obj.layoutSource,
					...(obj.definitions?.layout ? { layoutXml: obj.definitions.layout } : {}),
					...(obj.definitions?.colors ? { colorsXml: obj.definitions.colors } : {}),
					...(obj.definitions?.quickStyle ? { quickStyleXml: obj.definitions.quickStyle } : {}),
				},
				scheme: smartArtSchemeColors(theme),
				minorFont: theme?.minorFont,
				formatScheme: theme?.formatScheme,
			})
		: undefined;
	const rawDrawing = cached ?? preview?.drawing;
	const diagram = rawDrawing
		? resolveDiagramDrawingStyles(rawDrawing, theme?.formatScheme)
		: undefined;
	node.dataset.rendering = cached ? 'cached-drawing' : (preview?.mode ?? 'placeholder');
	if (preview?.issues.length)
		node.title = [notice, ...preview.issues.map((issue) => `${issue.code}: ${issue.message}`)].join(
			'\n',
		);
	if (diagram && diagram.shapes.length) {
		defineSmartArt(doc.defaultView?.customElements);
		const art = doc.createElement('office-ui-smartart') as HTMLElement & {
			drawing?: SmartArtObject['diagram'];
			schemeColors?: SchemeColors;
			schemeFonts?: SchemeFonts;
		};
		art.setAttribute('label', label);
		art.style.width = '100%';
		art.style.height = '100%';
		art.schemeColors = smartArtSchemeColors(theme);
		art.schemeFonts = theme ? { major: theme.majorFont, minor: theme.minorFont } : {};
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
	if (cached) return;
	const badge = h(doc, 'div', 'xg-smartart-notice');
	badge.textContent = ctx.t('Drawn with an approximate layout computed here; not Excel-identical.');
	badge.title = notice;
	node.append(badge);
}
