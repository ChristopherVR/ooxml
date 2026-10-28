import { NS, elements, type XmlElement } from '../xml/index';
export const ARTISTIC_PRIMARY_ATTRIBUTE: Readonly<Record<string, string>> = {
	artisticBlur: 'radius',
	artisticCement: 'crackSpacing',
	artisticChalkSketch: 'pressure',
	artisticCrisscrossEtching: 'pressure',
	artisticCutout: 'numberOfShades',
	artisticFilmGrain: 'grainSize',
	artisticGlass: 'scaling',
	artisticGlowDiffused: 'intensity',
	artisticGlowEdges: 'smoothness',
	artisticLightScreen: 'gridSize',
	artisticLineDrawing: 'pencilSize',
	artisticMarker: 'size',
	// Spelled "Mosiaic" in the schema; Microsoft's typo is part of the format.
	artisticMosiaicBubbles: 'pressure',
	artisticPaintBrush: 'brushSize',
	artisticPaintStrokes: 'intensity',
	artisticPastelsSmooth: 'scaling',
	artisticPencilGrayscale: 'pencilSize',
	artisticPencilSketch: 'pressure',
	artisticPhotocopy: 'detail',
	artisticPlasticWrap: 'smoothness',
	artisticTexturizer: 'scaling',
	artisticWatercolorSponge: 'brushSize',
};
export interface DrawingArtisticPictureEffect {
	kind: 'artistic';
	name: string;
	radius: number;
	baked: boolean;
	/** Extension children whose visual meaning is not understood. Source XML remains intact. */
	unhandled?: string[];
	sourceXml?: string;
}
/** Native imgProps record a destructively cached effect: never apply it a second time to the blip. */
export function parseArtisticPictureExtension(
	node: XmlElement,
): DrawingArtisticPictureEffect | undefined {
	if (node.namespaceURI !== NS.a || node.localName !== 'extLst') return undefined;
	const found: XmlElement[] = [];
	const unhandled: string[] = [];
	const pending = [node];
	while (pending.length) {
		const item = pending.pop()!;
		if (
			item !== node &&
			!(item.namespaceURI === NS.a && item.localName === 'ext') &&
			!(
				item.namespaceURI === 'http://schemas.microsoft.com/office/drawing/2010/main' &&
				(['imgProps', 'imgLayer', 'imgEffect', 'useLocalDpi'].includes(item.localName) ||
					ARTISTIC_PRIMARY_ATTRIBUTE[item.localName])
			)
		)
			unhandled.push(item.localName);
		if (
			item.namespaceURI === 'http://schemas.microsoft.com/office/drawing/2010/main' &&
			item.localName.startsWith('artistic')
		)
			found.push(item);
		pending.push(...elements(item));
	}
	if (found.length !== 1) return undefined;
	const effect = found[0]!;
	const key = ARTISTIC_PRIMARY_ATTRIBUTE[effect.localName];
	if (!key) return undefined;
	const raw = key ? effect.getAttribute(key) : null;
	const n = raw === null ? 0 : Number(raw);
	if (!Number.isFinite(n)) return undefined;
	return {
		kind: 'artistic',
		name: effect.localName,
		radius: effect.localName === 'artisticBlur' ? n : n / 1000,
		baked: true,
		...(unhandled.length ? { unhandled } : {}),
	};
}
