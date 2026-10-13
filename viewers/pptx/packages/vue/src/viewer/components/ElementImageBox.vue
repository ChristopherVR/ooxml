<script setup lang="ts">
/**
 * ElementImageBox: the picture/image branch of `ElementRenderer`, extracted to
 * keep the dispatcher thin. Renders an `<img>` under the shared fill/crop fit
 * with the computed CSS filter + any SVG `<filter>` defs for duotone/artistic
 * effects.
 */

import {
	getComputedImageStyle,
	getCropShapeClipPath,
	getImageColorWashStyle,
	getImageFillCounterTransform,
	getImageFitStyle,
	getImageOverflow,
	getImageTilingStyle,
	getPictureFrameLayers,
	resolveShapeGeometry,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxImageEffects } from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';
import type { CSSProperties } from 'vue';
import { computed } from 'vue';

import { useElementHitTargetStyle } from '../composables/element-hit-target';
import { getContainerStyle, getImageSrc } from '../composables/element-style';
import { useColorChangeImage } from '../composables/use-color-change-image';
import type { ClrChangeEffect } from '../composables/use-color-change-image';
import { useNativeImageSize } from '../composables/use-native-image-size';
import ShapeEffectOverlay from './ShapeEffectOverlay.vue';

const props = defineProps<{
	element: PptxElement;
	mediaDataUrls: Map<string, string>;
	zIndex: number;
	interactive?: boolean;
	/** Emit the data-pptx-element marker even when not interactive (template layer). */
	marked?: boolean;
	/** True only on the live presentation stage; see `hitTargetStyle`. */
	presenting?: boolean;
}>();

/**
 * The mask the picture's own shape geometry (`p:spPr/a:prstGeom` /
 * `a:custGeom`) imposes on the stationary container: `border-radius` for the
 * roundRect family and ellipse presets, a rescaled `clip-path` for custGeom
 * and other silhouettes. `undefined` for effectively rectangular pictures,
 * where the frame's overflow clipping already expresses the geometry. Kept
 * off the `<img>`, whose source-crop transform would scale and shift a
 * pixel-space clip.
 */
const geometryMask = computed<CSSProperties | undefined>(() => {
	if (!isImageLikeElement(props.element)) {
		return undefined;
	}
	const geometry = resolveShapeGeometry(props.element);
	if (geometry.kind === 'borderRadius') {
		return { borderRadius: geometry.radius };
	}
	return geometry.kind === 'clipPath' ? { clipPath: geometry.clipPath } : undefined;
});

// "Crop to Shape" (`element.cropShape`, the picture-format gallery, distinct
// from `<a:srcRect>` rectangular cropping above): a CSS `clip-path` on the
// stationary container. Shared's `getCropShapeClipPath` routes through the
// same adjustment-aware preset cascade every shape uses, so `roundedRect` and
// `star` render with their real geometry instead of a fixed approximation.
// Vue had no crop-to-shape support at all before this. The picture's own
// shape geometry outranks it; the crop shape is the fallback.
const cropShapeClip = computed<string | undefined>(() =>
	isImageLikeElement(props.element)
		? getCropShapeClipPath(props.element.cropShape, props.element.width, props.element.height)
		: undefined,
);

// The clip is load-bearing, not cosmetic: a cropped picture is rendered by
// scaling the source up and translating the cropped-away part out of the frame,
// so without it the discarded region paints over its neighbours. A picture with
// a frame (`a:ln`) or an outer shadow keeps that clip on the inner surface
// instead, so the centred line (`ShapeEffectOverlay`) and the shadow can extend
// past the box (shared `getPictureFrameLayers`).
const frameLayers = computed(() =>
	getPictureFrameLayers(props.element, {
		overflow: getImageOverflow(props.element),
		borderRadius: geometryMask.value?.borderRadius as string | undefined,
		clipPath: (geometryMask.value?.clipPath as string | undefined) ?? cropShapeClip.value,
	}),
);
const containerStyle = computed<CSSProperties>(() => ({
	...getContainerStyle(props.element, props.zIndex),
	...(frameLayers.value.box as CSSProperties),
}));
// Without a frame the surface takes no box of its own, so the markup lays out
// exactly as the single clipped box did.
const surfaceStyle = computed<CSSProperties>(
	() => (frameLayers.value.surface as CSSProperties | undefined) ?? { display: 'contents' },
);
const imageFitStyle = computed<CSSProperties>(
	() => getImageFitStyle(props.element) as CSSProperties,
);
// `a:blipFill/@rotWithShape="0"`: keep the picture's PIXELS upright/fixed to
// the page frame while the container above (which owns the geometry clip)
// still carries the shape's own rotate/flip. `undefined` (no-op) in the
// overwhelming majority of cases; see `getImageFillCounterTransform`.
const fillCounterTransform = computed<string | undefined>(() =>
	getImageFillCounterTransform(props.element, Boolean(imageFitStyle.value.transform)),
);
const imageSrc = computed(() => getImageSrc(props.element, props.mediaDataUrls));
// `a:blipFill/a:tile`: a repeating TEXTURE, which an `<img>` cannot express, so
// the picture paints as a repeating background layer instead. `undefined` for a
// normal (untiled) picture, which keeps the `<img>` branch.
//
// `@sx`/`@sy` is a percentage of the picture's own NATIVE pixel size (ECMA-376
// §20.1.8.58), not of this container, so `tileNativeSize` is probed (keyed on
// the same resolved `imageSrc`) and passed through: `getImageTilingStyle`
// falls back to the old container-relative percentage until it resolves.
const tileNativeSize = useNativeImageSize(imageSrc);
const tilingStyle = computed<CSSProperties | undefined>(
	() => getImageTilingStyle(props.element, tileNativeSize.value) as CSSProperties | undefined,
);
const imageFx = computed(() => getComputedImageStyle(props.element));
const colorWash = computed(() => {
	const effects = (props.element as { imageEffects?: PptxImageEffects }).imageEffects;
	return getImageColorWashStyle(effects?.colorWash);
});

// The `<a:clrChange>` chroma-key effect, if present (with a valid `clrFrom`).
const clrChange = computed<ClrChangeEffect | undefined>(() => {
	const effects = (props.element as { imageEffects?: PptxImageEffects }).imageEffects;
	const effect = effects?.clrChange;
	return effect?.clrFrom ? effect : undefined;
});

// Recoloured source (offscreen-canvas pixel swap); falls back to the original
// `imageSrc` while processing, on failure, or when no clrChange is present.
const { displaySrc } = useColorChangeImage({ src: imageSrc, clrChange });

/**
 * `pointer-events: none` while not interactive, mirroring React's
 * `pointer-events-none` class / Angular's `rootPointerEvents`. `marked` keeps
 * the element findable via `data-pptx-element` even while locked (e.g. a
 * template/master picture with `editTemplateMode` off); this is what actually
 * stops it from being clicked or dragged. `null` while interactive so the
 * style-array merge leaves any pre-existing `pointerEvents` untouched.
 */
const rootPointerEvents = computed<CSSProperties | null>(() =>
	props.interactive ? null : { pointerEvents: 'none' },
);

/**
 * Interaction-only affordance for a degenerate (sub-MIN_ELEMENT_SIZE) picture:
 * see `ElementRenderer`'s `hitTargetStyle` doc comment (issue #285). Never
 * rendered while presenting.
 */
const hitTargetStyle = useElementHitTargetStyle(
	() => props.element,
	() => props.interactive,
	() => props.presenting,
);
</script>

<template>
	<div
		class="pptx-vue-element pptx-vue-image"
		:style="[containerStyle, rootPointerEvents]"
		:data-element-id="element.id"
		:data-pptx-element="interactive || marked ? 'true' : undefined"
	>
		<!-- Interaction-only hit-target for a degenerate picture; see `hitTargetStyle`. -->
		<div
			v-if="hitTargetStyle"
			aria-hidden="true"
			data-pptx-hit-target="true"
			:style="hitTargetStyle"
		/>
		<div class="pptx-vue-image-surface" :style="surfaceStyle">
			<!-- SVG <filter> defs for duotone / advanced-alpha / artistic image effects. -->
			<svg
				v-for="f in imageFx.svgFilters"
				:key="f.id"
				width="0"
				height="0"
				aria-hidden="true"
				style="position: absolute; width: 0; height: 0; overflow: hidden"
			>
				<defs>
					<filter :id="f.id" color-interpolation-filters="sRGB" v-html="f.markup" />
				</defs>
			</svg>
			<div
				v-if="tilingStyle"
				class="pptx-vue-image-tile"
				:style="{ ...tilingStyle, filter: imageFx.filter, opacity: imageFx.opacity }"
			/>
			<!-- Not draggable: the browser's own image drag would cancel the pointer
		     stream of the editor's move gesture. -->
			<img
				v-else-if="imageSrc"
				:src="displaySrc ?? imageSrc"
				alt=""
				draggable="false"
				:style="{
					...imageFitStyle,
					display: 'block',
					filter: imageFx.filter,
					opacity: imageFx.opacity,
					...(fillCounterTransform
						? { transform: fillCounterTransform, transformOrigin: 'center' }
						: {}),
				}"
			/>
			<div
				v-if="colorWash"
				class="pptx-vue-image-color-wash"
				:style="{
					position: 'absolute',
					inset: '0',
					pointerEvents: 'none',
					backgroundColor: colorWash.backgroundColor,
					opacity: colorWash.opacity,
				}"
			/>
		</div>
		<!-- DAG fill-overlay tint, `a:reflection` mirror and the picture's own
		     frame (`a:ln`, the centred outline from `buildStrokeOutline`). It sits
		     beside the clipped surface, on the box, so the half of the line that
		     falls outside the picture is not cut. -->
		<ShapeEffectOverlay :element="element" :media-data-urls="mediaDataUrls" />
	</div>
</template>
