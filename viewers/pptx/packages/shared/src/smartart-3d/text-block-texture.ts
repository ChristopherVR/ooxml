/**
 * Canvas texture for a `SmartArt3DTextBlock`, for the `<pptx-three-view>`
 * SmartArt scene (`view-scene.ts`).
 *
 * Unlike the legacy `smartart-3d/text-texture.ts` (which imports `three`
 * statically and re-wraps text itself), this draws the ALREADY wrapped/
 * positioned lines a `SmartArt3DTextBlock` carries (built from the same 2D
 * projection the SVG renderer uses), and takes the runtime `three` module as
 * a parameter so it stays out of the tsup-bundled `dist/index.mjs` (see the
 * "no runtime `three` import" rule in `three-view/types.ts`).
 *
 * @module smartart-3d/text-block-texture
 */
import type * as THREE from 'three';

import type { SmartArt3DTextBlock } from '../render/smartart-3d-types';
import type { ThreeModule } from '../three-view/types';
import { layoutTextBlock } from './text-block-layout';

/** Anisotropic filtering for labels seen at a steep angle (three clamps it to the GPU's maximum). */
const ANISOTROPY = 8;

/** A built label texture plus the world-space plane size it should fill. */
export interface TextBlockTexture {
	texture: THREE.CanvasTexture;
	worldWidth: number;
	worldHeight: number;
}

/** Build a canvas texture rendering `block`'s pre-wrapped lines, or `null` off-DOM. */
export function buildTextBlockTexture(
	three: ThreeModule,
	block: SmartArt3DTextBlock,
): TextBlockTexture | null {
	if (typeof document === 'undefined' || block.lines.every((l) => l.text.length === 0)) {
		return null;
	}
	const canvas = document.createElement('canvas');
	const ctx2d = canvas.getContext('2d');
	if (!ctx2d) {
		return null;
	}

	const weight = block.fontWeight ?? 400;
	const italic = block.fontStyle === 'italic' ? 'italic ' : '';
	const family = block.fontFamily
		? `${block.fontFamily}, system-ui, sans-serif`
		: 'system-ui, -apple-system, Segoe UI, Roboto, sans-serif';
	const fontSize = Math.max(6, block.fontSize);
	const fontAt = (scale: number): string => `${italic}${weight} ${fontSize * scale}px ${family}`;

	// Measure at a fixed 4x so the layout does not depend on the canvas it sizes.
	const MEASURE = 4;
	ctx2d.font = fontAt(MEASURE);
	const layout = layoutTextBlock(
		block,
		(text) => ctx2d.measureText(text).width / MEASURE,
		typeof window === 'undefined' ? 1 : window.devicePixelRatio,
	);
	const { worldWidth, worldHeight, scale } = layout;
	// Resizing resets the context state, so every property is set after it.
	canvas.width = layout.pixelWidth;
	canvas.height = layout.pixelHeight;

	ctx2d.clearRect(0, 0, canvas.width, canvas.height);
	ctx2d.fillStyle = block.color;
	ctx2d.textAlign = 'center';
	// `dy` places each line's CENTRE (as `centeredSvgTextLines` lays it out,
	// and the 2D renderers draw with `dominant-baseline: central`).
	ctx2d.textBaseline = 'middle';
	ctx2d.font = fontAt(scale);
	const centerX = canvas.width / 2;
	const centerY = canvas.height / 2;
	for (const line of block.lines) {
		if (!line.text) {
			continue;
		}
		// `dy` is a y-up offset from the block centre; canvas y grows downward.
		ctx2d.fillText(line.text, centerX, centerY - line.dy * scale);
	}

	const texture = new three.CanvasTexture(canvas);
	texture.colorSpace = three.SRGBColorSpace;
	// Mipmaps and anisotropy keep a tilted or shrunk label smooth instead of shimmering.
	texture.generateMipmaps = true;
	texture.minFilter = three.LinearMipmapLinearFilter;
	texture.magFilter = three.LinearFilter;
	texture.anisotropy = ANISOTROPY;
	// See `smartart-3d/text-texture.ts` for why flipY must stay false with a
	// compensating UV flip: WebGL2 forbids UNPACK_FLIP_Y_WEBGL for some texture
	// targets, and leaving it enabled pollutes global pixel-store state.
	texture.flipY = false;
	texture.premultiplyAlpha = false;
	texture.wrapT = three.RepeatWrapping;
	texture.repeat.set(1, -1);
	texture.offset.set(0, 1);
	texture.needsUpdate = true;

	return { texture, worldWidth, worldHeight };
}
